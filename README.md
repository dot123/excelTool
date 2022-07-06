# excelTool
高性能高并发Excel导表工具支持导出.txt .lua .json

1.excel导表工具支持导出txt、lua、json文件。

2.支持格式参考测试例子.xlsx。

3.开头#号注释。

4.config.json可以配置导出路径如不填则不导出该文件。

5.fileList方便读取所有导出的文件。

6.支持数据类型有int、 float、 string、 auto。

6.配置说明

```
type Config struct {
	Root         string
	Txt          string
	JSON         string
	Lua          string
	FieldLine    int    //字段key开始行
	DataLine     int    //有效配置开始行
	TypeLine     int    //类型配置开始行
	Comma        string //txt分隔符,默认是制表符
	Comment      string //excel注释符
	Linefeed     string //txt换行符
	UseSheetName bool   //使用工作表名为文件输出名
}
```

#### 项目代码https://github.com/dot123/excelTool.git
