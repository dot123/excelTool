# excelTool
高性能高并发Excel导表工具支持导出txt、lua、json、msgpack格式json

1.excel导表工具支持导出txt、lua、json、msgpack格式json文件,并可以使用zlib压缩。

2.支持格式参考测试例子.xlsx。

3.开头#号注释。

4.表格列或行有留空则跳过。

5.config.json可以配置导出路径如不填则不导出该文件。

6.fileList方便读取所有导出的文件。

7.支持数据类型有int、float、string、auto。

8.auto类型自动推导类型如数字-128->int64,255->int64,12345.6->float32(float32 大约提供 6-7 位的精度),123456.7->float64。

9.配置说明

```
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
```

#### 项目代码https://github.com/dot123/excelTool.git
