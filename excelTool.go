package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io/ioutil"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
)

type config struct {
	Configs   string
	Txt       string
	JSON      string
	Lua       string
	FieldLine int //字段key开始行
	DataLine  int //有效配置开始行
}

var (
	ch         = make(chan string)
	fileCount  int
	configJSON config
	fileList   = make([]interface{}, 0)
)

func main() {
	c := flag.String("C", "./config.json", "配置文件路径")
	flag.Parse()

	fileCount = 0

	configJSON = config{}
	//读取json配置
	data, err := ioutil.ReadFile(*c)
	if err != nil {
		fmt.Printf("%v\n", err)
		return
	}

	err = json.Unmarshal(data, &configJSON)
	if err != nil {
		fmt.Printf("%v\n", err)
		return
	}

	//创建输出路径
	outList := [3]string{configJSON.Txt, configJSON.Lua, configJSON.JSON}
	for _, v := range outList {
		if v != "" {
			err = createDir(v)
			if err != nil {
				return
			}
		}
	}

	//遍历打印所有的文件名
	filepath.Walk(configJSON.Configs, walkFunc)
	count := 0
	for {
		fileName, open := <-ch
		if !open {
			break
		}
		fmt.Printf("已完成解析：%s\r\n", fileName+".xlsx")
		fileList = append(fileList, fileName) //添加到文件列表
		count++
		if count == fileCount {
			writeFileList()
			break
		}
	}
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

	if configJSON.Txt != "" {
		writeJSON(configJSON.Txt, "fileList", m)
	}

	if configJSON.JSON != "" {
		writeJSON(configJSON.JSON, "fileList", m)
	}

	if configJSON.Lua != "" {
		writeLuaTable(configJSON.Lua, "fileList", m)
	}
}

//创建文件夹
func createDir(dir string) error {
	exist, err := pathExists(dir)
	if err != nil {
		fmt.Printf("get dir error![%v]\n", err)
		return err
	}

	if exist {
		fmt.Printf("has dir![%v]\n", dir)
	} else {
		fmt.Printf("no dir![%v]\n", dir)
		//创建文件夹
		err := os.MkdirAll(dir, os.ModePerm)
		if err != nil {
			fmt.Printf("mkdir failed![%v]\n", err)
		} else {
			fmt.Printf("mkdir success!\n")
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
	if path.Ext(files) == ".xlsx" && !strings.HasPrefix(fileName, "~$") {
		fileCount++
		go readXlsx(files, strings.Replace(fileName, ".xlsx", "", -1))
	}
	return nil
}

//设置字段默认值
func setFieldDefault(fileName string, rowN int, rows []string, fieldCount int) []string {
	for i, row := range rows {
		if row == "" {
			// fmt.Printf("在%s中第%d行第%d个字段没有填，默认是0\n", fileName, rowN, i)
			// rows[i] = "0"
		}
		if i == fieldCount-1 {
			rows = append(rows[:i+1])
			break
		}
	}
	return rows
}

//检查一行是否有效
func checkRowValid(rows []string) bool {
	for _, row := range rows {
		if row != "" {
			return true
		}
	}
	return false
}

//读取xlsx
func readXlsx(path string, fileName string) {
	fmt.Println("正在解析：" + path)

	var file *os.File

	if configJSON.Txt != "" {
		_file, fileErr := os.OpenFile(configJSON.Txt+"/"+fileName+".txt", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666) //不存在创建清空内容覆写

		if fileErr != nil {
			fmt.Println("open file failed.", fileErr.Error())
			return
		}
		file = _file
		defer _file.Close() //关闭文件
	}

	//打开excel
	xlsx, err := excelize.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		return
	}
	// // Get value from cell by given worksheet name and axis.
	// cell := xlsx.GetCellValue("Sheet1", "B2")
	// fmt.Println(cell)
	// // Get all the rows in the Sheet1.

	SheetName := xlsx.GetSheetName(1)
	var rows, _ = xlsx.GetRows(SheetName)
	fields := rows[configJSON.FieldLine] //字段key
	lineNum := 0                         //行数
	dataDict := make(map[string]interface{})

	for i, field := range fields {
		if field == "" {
			fields = append(fields[:i])
			break
		}
	}

	fieldCount := len(fields)

	for rowN, row := range rows {
		lineData := make(map[string]interface{}) //一行数据
		fieldNum := 0

		//第几个字段
		if lineNum < configJSON.DataLine {
			lineNum++
			if file == nil { //不存在
				continue
			}

			if checkRowValid(row) == false {
				continue
			}
			row = setFieldDefault(fileName, rowN, row, fieldCount) //设置字段默认值

			for _, value := range row { //txt所有都要写
				// if strings.HasPrefix(row[0], "#") { //#号跳过
				// 	break
				// }

				file.WriteString(value + "\t")
			}
			file.WriteString("\r\n")
			continue
		}

		lineNum++

		if checkRowValid(row) == false {
			continue
		}

		if strings.HasPrefix(row[0], "#") { //#号跳过
			continue
		}

		row = setFieldDefault(fileName, rowN, row, fieldCount) //设置字段默认值

		for _, value := range row {

			if file != nil { //存在
				file.WriteString(value + "\t")
			}

			var m map[string]interface{}
			err = json.Unmarshal([]byte(value), &m) //尝试转换成map
			if err == nil {
				lineData[fields[fieldNum]] = m
				fieldNum++
				continue
			}

			var arr []interface{}
			err = json.Unmarshal([]byte(value), &arr) //尝试转换成数组

			if err == nil {
				lineData[fields[fieldNum]] = arr
			} else {
				f, err := strconv.ParseFloat(value, 64) //尝试转换为float64
				if err == nil {
					lineData[fields[fieldNum]] = f
				} else {
					lineData[fields[fieldNum]] = value
				}
			}
			fieldNum++
		}
		dataDict[row[0]] = lineData //第一个字段作为索引
		if file != nil {            //存在
			file.WriteString("\r\n")
		}
	}

	if configJSON.JSON != "" {
		writeJSON(configJSON.JSON, fileName, dataDict) //写JSON文件
	}

	if configJSON.Lua != "" {
		writeLuaTable(configJSON.Lua, fileName, dataDict) //写Lua文件
	}

	ch <- fileName
}

//字典转字符串
func map2Str(dataDict map[string]interface{}) string {
	mjson, err := json.Marshal(dataDict)
	if err != nil {
		fmt.Println(err)
		return ""
	}
	return string(mjson)
}

//写JSON文件
func writeJSON(path string, fileName string, dataDict map[string]interface{}) {
	file, err := os.OpenFile(path+"/"+fileName+".json", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0666) //不存在创建清空内容覆写
	if err != nil {
		fmt.Println("open file failed.", err.Error())
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
		fmt.Println("open file failed.", err.Error())
		return
	}

	defer file.Close()
	file.WriteString("return ")
	writeLuaTableContent(file, dataDict, 0)
}

//写Lua表内容
func writeLuaTableContent(fileHandle *os.File, data interface{}, idx int) {
	switch t := data.(type) {
	case float64:
		fileHandle.WriteString(fmt.Sprintf("%v", data)) //对于interface{}, %v会打印实际类型的值
	case string:
		fileHandle.WriteString(fmt.Sprintf(`"%v"`, data)) //对于interface{}, %v会打印实际类型的值
	case []interface{}:
		fileHandle.WriteString("{\n")
		a := data.([]interface{})
		for _, v := range a {
			addTabs(fileHandle, idx)
			writeLuaTableContent(fileHandle, v, idx+1)
			fileHandle.WriteString(",\n")
		}
		addTabs(fileHandle, idx-1)
		fileHandle.WriteString("}")
	case []string:
		fileHandle.WriteString("{\n")
		a := data.([]string)
		for _, v := range a {
			addTabs(fileHandle, idx)
			writeLuaTableContent(fileHandle, v, idx+1)
			fileHandle.WriteString(",\n")
		}
		addTabs(fileHandle, idx-1)
		fileHandle.WriteString("}")
	case map[string]interface{}:
		m := data.(map[string]interface{})
		fileHandle.WriteString("{\n")
		for k, v := range m {
			addTabs(fileHandle, idx)
			fileHandle.WriteString("[")
			writeLuaTableContent(fileHandle, k, idx+1)
			fileHandle.WriteString("] = ")
			writeLuaTableContent(fileHandle, v, idx+1)
			fileHandle.WriteString(",\n")
		}
		addTabs(fileHandle, idx-1)
		fileHandle.WriteString("}")
	default:
		fileHandle.WriteString(fmt.Sprintf("%t", data))
		_ = t
	}
}

//在文件中添加制表符
func addTabs(fileHandle *os.File, idx int) {
	for i := 0; i < idx; i++ {
		fileHandle.WriteString("\t")
	}
}
