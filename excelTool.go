package main

import (
	"bytes"
	"compress/zlib"
	"encoding/json"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/shamaton/msgpack/v2"
	log "github.com/sirupsen/logrus"
	"math"
	"os"
	"path"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

type Config struct {
	Root         string // excel配置表根目录
	Txt          string // txt格式导出路径
	JSON         string // json格式导出路径
	Lua          string // lua格式导出路径
	Bin          string // msgpack格式json导出路径
	FieldLine    int    // 字段key开始行
	TypeLine     int    // 类型配置开始行
	DataLine     int    // 有效配置开始行
	UseZlib      bool   // 是否使用zlib压缩
	Comma        string // txt分隔符,默认是制表符
	Comment      string // excel注释符
	Linefeed     string // txt换行符
	UseSheetName bool   // 使用工作表名为文件输出名
}

var (
	config   Config
	fileList []interface{}
	wg       sync.WaitGroup
	rwMutex  sync.RWMutex
)

func main() {
	log.SetFormatter(&log.TextFormatter{ForceColors: true, FullTimestamp: true})

	startTime := time.Now().UnixNano()

	c := flag.String("C", "./config.json", "配置文件路径")
	flag.Parse()

	// 读取json配置
	data, err := os.ReadFile(*c)
	if err != nil {
		log.Fatal(err)
	}

	if err = json.Unmarshal(data, &config); err != nil {
		log.Fatal(err)
	}

	// 创建输出路径
	outList := []string{config.Txt, config.Lua, config.JSON, config.Bin}
	for _, v := range outList {
		if v != "" {
			err = createDir(v)
			if err != nil {
				return
			}
		}
	}

	// 遍历打印所有的文件名
	if err := filepath.Walk(config.Root, walkFunc); err != nil {
		log.Fatal(err)
	}

	wg.Wait()

	writeFileList()

	endTime := time.Now().UnixNano()
	log.Infof("总耗时:%v毫秒\n", (endTime-startTime)/1000000)
	time.Sleep(time.Second)
}

// 写文件列表
func writeFileList() {
	data := make(map[string]interface{})

	sortList := make([]string, len(fileList))
	for i, v := range fileList {
		sortList[i] = v.(string)
	}
	sort.Strings(sortList)
	data["fileList"] = sortList

	if config.Txt != "" {
		writeJSON(config.Txt, "fileList", &data)
	}

	if config.JSON != "" {
		writeJSON(config.JSON, "fileList", &data)
	}

	if config.Lua != "" {
		writeLuaTable(config.Lua, "fileList", &data)
	}

	if config.Bin != "" {
		writeBin(config.Bin, "fileList", &data)
	}
}

// 创建文件夹
func createDir(dir string) error {
	exist, err := pathExists(dir)
	if err != nil {
		log.Fatalf("get dir error![%v]\n", err)
		return err
	}

	if !exist {
		if err := os.MkdirAll(dir, os.ModePerm); err != nil {
			log.Fatalf("mkdir failed![%v]\n", err)
		} else {
			log.Infof("mkdir success!\n")
		}
	}
	return nil
}

// 判断文件夹是否存在
func pathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

func walkFunc(files string, info os.FileInfo, err error) error {
	_, fileName := filepath.Split(files)
	if path.Ext(files) == ".xlsx" && !strings.HasPrefix(fileName, "~$") && !strings.HasPrefix(fileName, "#") {
		wg.Add(1)
		go parseXlsx(files, strings.Replace(fileName, ".xlsx", "", -1))
	}
	return nil
}

// 解析xlsx
func parseXlsx(path string, fileName string) {
	defer wg.Done()

	// 打开excel
	xlsx, err := excelize.OpenFile(path)
	if err != nil {
		log.Errorf("%s %s", fileName, err)
		return
	}

	sheetName := xlsx.GetSheetName(1)
	var lines = xlsx.GetRows(sheetName)

	fieldMap := map[int]string{}
	set := map[string]int{}
	fieldList := lines[config.FieldLine-1]
	for i, field := range fieldList {
		if idx, ok := set[field]; ok {
			fieldMap[i] = fieldMap[idx]
			delete(fieldMap, idx)
		} else {
			if field != "" {
				set[field] = i
				fieldMap[i] = field
			}
		}
	}

	var idxList []int
	for k := range fieldMap {
		idxList = append(idxList, k)
	}

	sort.Ints(idxList)
	if len(idxList) == 0 {
		return
	}
	lineStart := idxList[0] // 主键第几列

	fieldCount := len(idxList)

	var fields []string
	for _, idx := range idxList {
		fields = append(fields, fieldMap[idx])
	}

	var data []interface{}
	data = append(data, fields)
	var buffer bytes.Buffer

	totalLineNum := len(lines)
	for n, line := range lines {
		if len(line) == 0 {
			continue
		}

		if strings.HasPrefix(line[0], config.Comment) { // 注释符跳过
			continue
		}

		if line[lineStart] == "" { // 主键不能为空
			log.Errorf("%s.xlsx (row=%v,col=%d) error: is '主键不能为空' \n", sheetName, n+1, lineStart+1)
			continue
		}

		if config.Txt != "" {
			fieldNum := 0
			for i, value := range line {
				if _, ok := fieldMap[i]; ok {
					fieldNum++
					buffer.WriteString(value)
					if fieldNum < fieldCount {
						buffer.WriteString(config.Comma)
					}
				}
			}
			if n < totalLineNum {
				buffer.WriteString(config.Linefeed)
			}
		}

		if n < config.DataLine-1 {
			continue
		}

		var lineData []interface{}
		for i, value := range line {
			if _, ok := fieldMap[i]; ok {
				lineData = append(lineData, typeConvert(lines[config.TypeLine-1][i], value))
			}
		}
		data = append(data, lineData)
	}

	if !config.UseSheetName {
		sheetName = fileName
	}

	if config.Txt != "" {
		writeTxt(config.Txt, sheetName, &buffer)
	}

	if config.JSON != "" {
		writeJSON(config.JSON, sheetName, &data)
	}

	if config.Lua != "" {
		writeLuaTable(config.Lua, sheetName, &data)
	}

	if config.Bin != "" {
		writeBin(config.Bin, sheetName, &data)
	}

	rwMutex.Lock()
	fileList = append(fileList, sheetName)
	rwMutex.Unlock()
}

func toFloat(value string) interface{} {
	float64Value, err := strconv.ParseFloat(value, 64)
	if err == nil {
		if float64Value >= math.SmallestNonzeroFloat32 && float64Value <= math.MaxFloat32 {
			// float32 大约提供 6-7 位的精度
			if len(value) <= 7 {
				return float32(float64Value)
			}
		}
		return float64Value
	}
	return nil
}

func toInt(value string, force bool) interface{} {
	if force {
		value = strings.Split(value, ".")[0]
	}

	intValue, err := strconv.ParseInt(value, 10, 64)
	if err != nil {
		return nil
	}
	return intValue
}

func toNumber(value string) interface{} {
	if uintValue := toInt(value, false); uintValue != nil {
		return uintValue
	}

	if floatValue := toFloat(value); floatValue != nil {
		return floatValue
	}

	return nil
}

func processNestedJSON(jsonData interface{}, key interface{}, value interface{}) {
	switch v := value.(type) {
	case map[string]interface{}:
		// 处理嵌套的 JSON 对象
		for nestedKey, nestedValue := range v {
			processNestedJSON(v, nestedKey, nestedValue)
		}
	case []interface{}:
		// 处理嵌套的 JSON 对象
		for nestedKey, nestedValue := range v {
			processNestedJSON(v, nestedKey, nestedValue)
		}
	case float64:
		str := strconv.FormatFloat(v, 'f', -1, 64)
		if num := toNumber(str); num != nil {
			switch m := jsonData.(type) {
			case map[string]interface{}:
				m[key.(string)] = num
			case []interface{}:
				m[key.(int)] = num
			}
		} else {
			log.Fatal(fmt.Errorf("failed to convert '%s' to a valid number", value))
		}
	}
}

func toJson(data []byte) interface{} {
	var jsonData map[string]interface{}
	if err := json.Unmarshal(data, &jsonData); err == nil {
		for key, value := range jsonData {
			processNestedJSON(jsonData, key, value)
		}
		return jsonData
	}

	var arr []interface{}
	if err := json.Unmarshal(data, &arr); err == nil {
		for key, value := range arr {
			processNestedJSON(arr, key, value)
		}
		return arr
	}

	return nil
}

// 类型转换
func typeConvert(ty string, value string) interface{} {
	switch ty {
	case "int":
		intValue := toInt(value, true)
		if intValue != nil {
			return intValue
		}
	case "float":
		floatValue := toFloat(value)
		if floatValue != nil {
			return floatValue
		}
	case "string":
		return value
	case "auto":
		m := toJson([]byte(value))
		if m != nil {
			return m
		}
	default:
		log.Fatalf("error in type %s\n", ty)
	}

	log.Fatal(fmt.Errorf("failed to convert '%s' to a valid number", value))

	return nil
}

// 写txt文件
func writeTxt(path string, fileName string, buffer *bytes.Buffer) {
	writeToFile(buffer.Bytes(), path+"/"+fileName+".txt")
}

// 写JSON文件
func writeJSON(path string, fileName string, data interface{}) {
	b, err := json.Marshal(data)
	if err != nil {
		log.Fatal(err)
	}

	writeToFile(b, path+"/"+fileName+".json")
}

// 写Lua文件
func writeLuaTable(path string, fileName string, data interface{}) {
	var buffer bytes.Buffer
	buffer.WriteString("return ")
	writeLuaTableContent(&buffer, data, 0)

	writeToFile(buffer.Bytes(), path+"/"+fileName+".lua")
}

// 写Lua表内容
func writeLuaTableContent(buffer *bytes.Buffer, data interface{}, idx int) {
	if data == nil {
		buffer.WriteString("nil")
		return
	}

	// 如果是指针类型
	if reflect.ValueOf(data).Type().Kind() == reflect.Pointer {
		data = reflect.ValueOf(data).Elem().Interface()
	}

	switch t := data.(type) {
	case int8, uint8, int16, uint16, int32, uint32, int64, uint64:
		buffer.WriteString(fmt.Sprintf("%d", data))
	case float32, float64:
		buffer.WriteString(fmt.Sprintf("%v", data))
	case string:
		buffer.WriteString(fmt.Sprintf(`"%s"`, data))
	case []interface{}:
		buffer.WriteString("{\n")
		a := data.([]interface{})
		for _, v := range a {
			addTabs(buffer, idx)
			writeLuaTableContent(buffer, v, idx+1)
			buffer.WriteString(",\n")
		}
		addTabs(buffer, idx-1)
		buffer.WriteString("}")
	case []string:
		buffer.WriteString("{\n")
		a := data.([]string)
		sort.Strings(a)
		for _, v := range a {
			addTabs(buffer, idx)
			writeLuaTableContent(buffer, v, idx+1)
			buffer.WriteString(",\n")
		}
		addTabs(buffer, idx-1)
		buffer.WriteString("}")
	case map[string]interface{}:
		m := data.(map[string]interface{})
		var keys []string
		for k := range m {
			keys = append(keys, k)
		}
		sort.Strings(keys)

		buffer.WriteString("{\n")
		for _, k := range keys {
			addTabs(buffer, idx)
			buffer.WriteString("[")
			writeLuaTableContent(buffer, k, idx+1)
			buffer.WriteString("] = ")
			writeLuaTableContent(buffer, m[k], idx+1)
			buffer.WriteString(",\n")
		}
		addTabs(buffer, idx-1)
		buffer.WriteString("}")
	default:
		buffer.WriteString(fmt.Sprintf("%t", data))
		_ = t
	}
}

// 在文件中添加制表符
func addTabs(buffer *bytes.Buffer, idx int) {
	for i := 0; i < idx; i++ {
		buffer.WriteString("\t")
	}
}

// 写bin文件
func writeBin(path string, fileName string, data interface{}) {
	b, err := msgpack.Marshal(data)
	if err != nil {
		log.Fatal(err)
	}

	writeToFile(b, path+"/"+fileName+".bin")
}

func writeToFile(inputData []byte, path string) {
	file, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666)
	if err != nil {
		log.Fatal("open file failed.", err.Error())
	}

	defer file.Close()

	if config.UseZlib {
		if _, err := file.Write(compressData(inputData)); err != nil {
			log.Fatal(err)
		}
	} else {
		if _, err := file.Write(inputData); err != nil {
			log.Fatal(err)
		}
	}
}

func compressData(inputData []byte) []byte {
	var compressedBuffer bytes.Buffer
	compressor := zlib.NewWriter(&compressedBuffer)

	if _, err := compressor.Write(inputData); err != nil {
		log.Fatal(err)
	}

	if err := compressor.Close(); err != nil {
		log.Fatal(err)
	}

	return compressedBuffer.Bytes()
}
