
# sun-wordtable-read
========
读取Word文档的各种复杂表格内容，支持2007以上的Docx文档（暂不支持2007以下的Doc类型文档）

## 背景：
工作上遇到如何读取Word文档中的表格内容，表格是有业务数据意义的，而且有一定规则的，因此不能直接读取表格文本，而是遍历表格单元格进行一行一列读取。

表格规则：
 1. 表格可以有表头，表头也有业务意思
 2. 一行为一个业务数据，可能会跨行
 3. 列可能会有跨列、跨行
 4. 单元格中图片、数学公式、嵌套表格、文件等

比如，以下表格
[![](https://img-blog.csdn.net/20180414152908387?watermark/2/text/aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3N1bmN0/font/5a6L5L2T/fontsize/400/fill/I0JBQkFCMA==/dissolve/70)](https://img-blog.csdn.net/20180414152908387?watermark/2/text/aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3N1bmN0/font/5a6L5L2T/fontsize/400/fill/I0JBQkFCMA==/dissolve/70)

## 设计理念：
 1. 读取Word文档中表格数据到内存映射表，再通过自定义读取策略，将内存映射表转换成实际业务表格数据。
 2. 使用统一的内存映射表，屏蔽了实际Word文档读取方式，开发者只关心如何转换为业务数据。

## 现状：
 1. 目前只支持读取2007以上Word文档表格单元格的文本，支持读取图片、数学公式、嵌套表格。
 2. 支持一般性的有规则的复杂表格。
 3. 暂不支持2007以下的Doc类型文档，因为POI中暂未找到关于表格单元格合并信息的API。
 4. 为了兼容2007以下的Doc类型文档，利用jodconverter3.0 + LibreOffice 5.3，“先将Doc类型文档转换为Docx类型文档，再进行读取表格内容”。
	注意：LibreOffice直接支持Docx类型文档，而OpenOffice不能直接支持Docx类型文档，需要AccessODF插件

## 后续要增加的功能：
 1. 读取单元格中其他附件
 2. 直接导入到目标（比如：数据库表、Excel等）的公共功能
 3. 读取大文件的Word、性能优化策略