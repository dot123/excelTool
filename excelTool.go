package main

import (
	"bytes"
	"encoding/json"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	log "github.com/sirupsen/logrus"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"time"
)

type Config struct {
	Configs      string
	Txt          string
	JSON         string
	Lua          string
	FieldLine    int    //字段key开始行
	DataLine     int    //有效配置开始行
	Comma        string //txt分隔符,默认是制表符
	Comment      string //excel注释符
	Linefeed     string //txt换行符
	UseSheetName bool   //使用工作表名为文件输出名
}

var (
	ch        = make(chan string)
	fileCount int
	config    Config
	fileList  = make([]interface{}, 0)
)

func main() {
	log.SetFormatter(&log.TextFormatter{ForceColors: true, FullTimestamp: true})

	startTime := time.Now().UnixNano()

	c := flag.String("C", "./config.json", "配置文件路径")
	flag.Parse()

	fileCount = 0

	config = Config{}
	//读取json配置
	data, err := ioutil.ReadFile(*c)
	if err != nil {
		log.Fatalf("%v\n", err)
		return
	}

	err = json.Unmarshal(data, &config)
	if err != nil {
		log.Fatalf("%v\n", err)
		return
	}

	//创建输出路径
	outList := [3]string{config.Txt, config.Lua, config.JSON}
	for _, v := range outList {
		if v != "" {
			err = createDir(v)
			if err != nil {
				return
			}
		}
	}

	//遍历打印所有的文件名
	filepath.Walk(config.Configs, walkFunc)
	count := 0
	for {
		sheetName, open := <-ch
		if !open {
			break
		}

		if sheetName != "" {
			fileList = append(fileList, sheetName) //添加到文件列表
		}

		count++
		if count == fileCount {
			writeFileList()
			break
		}
	}

	endTime := time.Now().UnixNano()
	log.Infof("总耗时:%v毫秒\n", (endTime-startTime)/1000000)
	time.Sleep(time.Millisecond * 1500)
}

//写文件列表
func writeFileList() {
	m := make(map[string]interface{}, 1)

	sortList := make([]string, len(fileList))
	for i, v := range fileList {
		sortList[i] = v.(string)
	}
	sort.Strings(sortList)
	m["fileList"] = sortList

	if config.Txt != "" {
		writeJSON(config.Txt, "fileList", m)
	}

	if config.JSON != "" {
		writeJSON(config.JSON, "fileList", m)
	}

	if config.Lua != "" {
		writeLuaTable(config.Lua, "fileList", m)
	}
}

//创建文件夹
func createDir(dir string) error {
	exist, err := pathExists(dir)
	if err != nil {
		log.Fatalf("get dir error![%v]\n", err)
		return err
	}

	if exist {
		log.Infof("has dir![%v]\n", dir)
	} else {
		log.Infof("no dir![%v]\n", dir)
		//创建文件夹
		err := os.MkdirAll(dir, os.ModePerm)
		if err != nil {
			log.Fatalf("mkdir failed![%v]\n", err)
		} else {
			log.Infof("mkdir success!\n")
		}
	}
	return nil
}

//判断文件夹是否存在
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
	// fmt.Println(paths, fileName)      //获取路径中的目录及文件名
	// fmt.Println(filepath.Base(files)) //获取路径中的文件名
	// fmt.Println(path.Ext(files))      //获取路径中的文件的后缀
	if path.Ext(files) == ".xlsx" && !strings.HasPrefix(fileName, "~$") && !strings.HasPrefix(fileName, "#") {
		fileCount++
		go readXlsx(files, strings.Replace(fileName, ".xlsx", "", -1))
	}
	return nil
}

//读取xlsx
func readXlsx(path string, fileName string) {
	//打开excel
	xlsx, err := excelize.OpenFile(path)
	if err != nil {
		log.Errorf("%s %s", fileName, err)
		ch <- ""
		return
	}
	// // Get value from cell by given worksheet name and axis.
	// cell := xlsx.GetCellValue("Sheet1", "B2")
	// fmt.Println(cell)
	// // Get all the rows in the Sheet1.

	var buffer bytes.Buffer

	sheetName := xlsx.GetSheetName(1)
	var lines = xlsx.GetRows(sheetName)

	fields := lines[config.FieldLine-1] //字段key
	lineNum := 0                        //行数
	dataDict := make(map[string]interface{})

	for i, field := range fields {
		if field == "" {
			fields = append(fields[:i])
			break
		}
	}

	fieldCount := len(fields)
	totalLineNum := len(lines)

	for n, line := range lines {
		data := make(map[string]interface{}) //一行数据
		fieldNum := 0

		if strings.HasPrefix(line[0], config.Comment) { //注释符跳过
			continue
		}

		if line[0] == "" {
			log.Errorf("%s.xlsx (row=%v,col=0) error: is '' \n", fileName, n+1)
			continue
		}
		line = line[0:fieldCount]

		lineNum++
		//第几个字段
		if lineNum < config.DataLine {
			for _, value := range line { //txt所有都要写
				fieldNum++
				buffer.WriteString(value)
				if fieldNum < fieldCount {
					buffer.WriteString(config.Comma)
				}
			}
			if lineNum < totalLineNum {
				buffer.WriteString(config.Linefeed)
			}
			continue
		}

		for _, value := range line {
			key := fields[fieldNum]
			fieldNum++
			buffer.WriteString(value)
			if fieldNum < fieldCount {
				buffer.WriteString(config.Comma)
			}

			var m map[string]interface{}
			err = json.Unmarshal([]byte(value), &m) //尝试转换成map
			if err == nil {
				data[key] = m
				continue
			}

			var arr []interface{}
			err = json.Unmarshal([]byte(value), &arr) //尝试转换成数组

			if err == nil {
				data[key] = arr
			} else {
				f, err := strconv.ParseFloat(value, 64) //尝试转换为float64
				if err == nil {
					data[key] = f
				} else {
					data[key] = value
				}
			}
		}
		dataDict[line[0]] = data //第一个字段作为索引

		if lineNum < totalLineNum {
			buffer.WriteString(config.Linefeed)
		}
	}

	if !config.UseSheetName {
		sheetName = fileName
	}

	if config.Txt != "" {
		writeTxt(config.Txt, sheetName, &buffer) //写txt文件
	}

	if config.JSON != "" {
		writeJSON(config.JSON, sheetName, dataDict) //写JSON文件
	}

	if config.Lua != "" {
		writeLuaTable(config.Lua, sheetName, dataDict) //写Lua文件
	}

	ch <- sheetName
}

//字典转字符串
func map2Str(dataDict map[string]interface{}) string {
	b, err := json.Marshal(dataDict)
	if err != nil {
		log.Errorln(err)
		return ""
	}
	return string(b)
}

//写txt文件
func writeTxt(path string, fileName string, buffer *bytes.Buffer) {
	file, err := os.OpenFile(path+"/"+fileName+".txt", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666) //不存在创建清空内容覆写
	if err != nil {
		log.Errorln("open file failed. ", err.Error())
		return
	}
	defer file.Close()
	file.Write(buffer.Bytes())
}

//写JSON文件
func writeJSON(path string, fileName string, dataDict map[string]interface{}) {
	file, err := os.OpenFile(path+"/"+fileName+".json", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666) //不存在创建清空内容覆写
	if err != nil {
		log.Errorln("open file failed.", err.Error())
		return
	}

	defer file.Close()
	//字典转字符串
	file.WriteString(map2Str(dataDict))
}

//写Lua文件
func writeLuaTable(path string, fileName string, dataDict interface{}) {
	file, err := os.OpenFile(path+"/"+fileName+".lua", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666) //不存在创建清空内容覆写
	if err != nil {
		log.Errorln("open file failed.", err.Error())
		return
	}

	defer file.Close()
	file.WriteString("return ")
	writeLuaTableContent(file, dataDict, 0)
}

//写Lua表内容
func writeLuaTableContent(file *os.File, data interface{}, idx int) {
	switch t := data.(type) {
	case float64:
		file.WriteString(fmt.Sprintf("%v", data)) //对于interface{}, %v会打印实际类型的值
	case string:
		file.WriteString(fmt.Sprintf(`"%v"`, data)) //对于interface{}, %v会打印实际类型的值
	case []interface{}:
		file.WriteString("{\n")
		a := data.([]interface{})
		for _, v := range a {
			addTabs(file, idx)
			writeLuaTableContent(file, v, idx+1)
			file.WriteString(",\n")
		}
		addTabs(file, idx-1)
		file.WriteString("}")
	case []string:
		file.WriteString("{\n")
		a := data.([]string)
		for _, v := range a {
			addTabs(file, idx)
			writeLuaTableContent(file, v, idx+1)
			file.WriteString(",\n")
		}
		addTabs(file, idx-1)
		file.WriteString("}")
	case map[string]interface{}:
		m := data.(map[string]interface{})
		file.WriteString("{\n")
		for k, v := range m {
			addTabs(file, idx)
			file.WriteString("[")
			writeLuaTableContent(file, k, idx+1)
			file.WriteString("] = ")
			writeLuaTableContent(file, v, idx+1)
			file.WriteString(",\n")
		}
		addTabs(file, idx-1)
		file.WriteString("}")
	default:
		file.WriteString(fmt.Sprintf("%t", data))
		_ = t
	}
}

//在文件中添加制表符
func addTabs(file *os.File, idx int) {
	for i := 0; i < idx; i++ {
		file.WriteString("\t")
	}
}
