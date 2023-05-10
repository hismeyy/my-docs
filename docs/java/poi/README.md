# 1. 前言

在开发中进行遇到Excel的导入和导出操作。目前操作Excel的框架有两个：

- apache的poi

  Apache POI是用Java编写的免费开源的跨平台的 Java API，Apache POI提供API给Java程式对Microsoft Office（Excel、WORD、PowerPoint、Visio)等）格式档案读和写的功能。

- Java Excel

  Java Excel是一开放源码项目，通过它Java开发人员可以读取Excel文件的内容、创建新的Excel文件、更新已经存在的Excel文件。jxl 由于其小巧 易用的特点, 逐渐已经取代了 POI-excel的地位, 成为了越来越多的java开发人员生成excel文件的首选。

该文章主要介绍POI

# 2. POI相关资料

> 官网：https://poi.apache.org/index.html
>
> 参考文章：https://blog.csdn.net/hadues/article/details/113859228

# 第一部分：Excel文件处理组件

# 1. 处理Excel文件

|                            POIFS                             |     HSSF      |      XSSF      |      SXSSF       |
| :----------------------------------------------------------: | :-----------: | :------------: | :--------------: |
| OIFS是POI中最古老，最稳定的部分。OLE 2复合文档格式到纯Java的移植 | 读写*.xls文件 | 读写*.xlsx文件 | 读写*.xlsx大文件 |

# 2. Maven依赖

```xml
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.0.0</version>
</dependency>
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->
<dependency>
    <groupId>org.apache.poi</groupId>
	<artifactId>poi-ooxml</artifactId>
	<version>5.0.0</version>
</dependency>
```

> 点这里查看最新版：https://mvnrepository.com/artifact/org.apache.poi/poi

# 3. 使用POI

## 3.1 创建WorkBook

1. 创建组件对象

   ```java
   XSSFWorkbook xwb = new XSSFWorkbook();
   ```

2. 创建文件输出流

   ```java
   FileOutputStream xfile = new FileOutputStream("xtest01.xlsx");
   ```

3. 写出文件

   ```java
    xwb.write(xfile);
   ```


## 3.2 创建Sheet

1. 创建WorkBook对象

2. 创建XSSFSheet对象

   ```java
   XSSFSheet sheet = xwb.createSheet("第一个文档");
   ```

3. 写出文件

## 3.3 创建单元格

1. 创建WorkBook对象

2. 创建Sheet对象

3. 创建Row对象

   ```java
   XSSFRow row = sheet.createRow(0);
   ```

4. 创建Cell对象

   ```java
   XSSFCell cell = row.createCell(0);
   ```

5. 设置单元格内容等

## 3.4 样式设置

1. 设置列宽

   需要用Sheet对象```setColumnWidth()```方法 width：1000 = 3.25

2. 设置行高

   需要用Row对象```setHeight()```方法 height：1000 = 50

3. 设置单元格样式

   需要使用WorkBook对象创建一个样式对象```XSSFCellStyle```，再设置给Cell对象
