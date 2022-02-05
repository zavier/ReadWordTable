# ReadWordTable

使用poi解析word文档（.docx）中的表格内容及结构，并以html形式输出

原理分析：https://www.cnblogs.com/zawier/p/6062596.html

总结一下处理Word中表格的行列合并相关的api:

获取单元格的属性

```java
XWPFTableCell cell = table.getRow(2).getCell(0);
CTTcPr tcPr = cell.getCTTc().getTcPr();
```

行合并时

```java
CTVMerge vMerge = tcPr.getVMerge();
// 如果不是行合并的单元格，则没有此属性，即： vMerge == null
// 如果是行合并的第一行单元格，则： vMerge.getVal().toString() == "restart"
// 如果是行合并的其他行单元格，则： vMerge.getVal() 
```

列合并时

```java
CTDecimalNumber gridSpan = tcPr.getGridSpan();
// 其他不是列合并的单元格，则：gridSpan == null
// 如果是列合并的第一列单元格，则：gridSpan.getVal()可以获取到这列单元格所占的行数
```





