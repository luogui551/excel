# 基于poi3.17对excel操作的简单封装
包含简单的读取、编辑功能，不包含所有的样式相关操作
## 获取对象
Excel excel = ExcelUtil.parse(excelFile)/ExcelUtil.newExcel(sheetName);//或者调用构造方法，重载方法，支持文件或流
## 遍历 excel.iterate/iterateAll
### 遍历当前sheet
do {
	excel.iterate...
}while(excel.nextSheet())
### 遍历所有sheet
excel.iterateAll

## 编辑
### 创建新sheet页
excel.newSheet(sheetName);
### 设置宽度
excel.setWidths
### 写数据
excel.write/writeRow
### 合并单元格
excel.merge

#...
