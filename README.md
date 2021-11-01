# POILIGHT

[![Build Status](https://travis-ci.org/lynchmaniac/poilight.svg?branch=master)](https://travis-ci.org/lynchmaniac/poilight)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Poilight is a wrapper of POI to accelerate the generation of Excel's files. Poilight only deals with Excel's files.

## Installation

Poilight is a Maven artefact so you can put the dependency below in your pom.xml to use Poilight in your project.

```java
    <dependencies>
        <dependency>
            <groupId>com.github.lynchmaniac</groupId>
            <artifactId>poilight</artifactId>
            <version>0.1.2</version>
        </dependency>
    </dependencies>
```

## Generate a simple File

The principle of poilight is to simply manage Excel spreadsheets. It's simple tables corresponding to the output of your Java applications. You must set a complete full path for your file and fill some data. For that, you can use the Table object, see the next chapter.

```java
    String outputPathExcelFile = "file.xlsx";
    Table table = new Table();
    ...
    PoiLight.generateExcel(outputPathExcelFile, table);
```

## Structure of the data

The data is a Table object. It is the table you want to achieve. If you want to add a header, you can specify them with the addHeader method. It takes a ExcelCell. ExcelCell is a structure in which you can precise the value and the of the cell. We will see this in a later chapter.
You can pass data to the table object with the method addData which takes a ExcelRow as parameter. ExcelRow is just a list of ExcelCell objects.
Below is an example of a basic table :

```java
    Table table = new Table();
    table.addHeaders("ID", "NOM", "TITRE");
    table.addData(new ExcelRow(1, "Henri Loevenbruck", "L'apothicaire"));
    table.addData(new ExcelRow(2, "Cyril Massarotto", "Dieu est un pote à moi"));
    table.addData(new ExcelRow(3, "Bernard Werber", "Les fourmis"));
    table.addData(new ExcelRow(4, "Maxime Chattam", "In Tenebris"));
    table.addData(new ExcelRow(5, "Franck Thilliez", "Pandemia"));

```

In poilight for a table, each row is represent by an ExcelRow object and each cell is represent by an ExcelCell object.

### Specific style

If you want to add a specific style to a cell, you can replace the strings of the header o data by an ExcelCell object. This accept in second parameter a CellStyle.
You can customize all the style you want with this object.

```java
    Table table = new Table();
    table.addHeaders("ID", "NOM", "TITRE");
    CellStyle cs = wb.createCellStyle();
    cs.setFillForegroundColor(StyleHelper.getColor(128, 100, 162).getIndex());
    cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
    ....
    table.addData(new ExcelRow(new ExcelCell(4, cs), "Maxime Chattam", "In Tenebris"));
    table.addData(new ExcelRow(5, "Franck Thilliez", "Pandemia"));
```

Here we have a specific style on the first cell. The background is set to purple. With this mechanism you can customize style as you like.

### The predefined styles

If you don't want to specify a style, Poilight embeds the entire 60 preset styles in Excel just for you. You can specify a style with the BoardStyles enum.

```java
    Table table = new Table();
    table.setStyle(BoardStyles.BOARD_LIGHT_RED_3_STYLE);
    PoiLight.generateExcel(outputPathExcelFile, table);
```

You can find in the example directory an Excel File with the 60 styles.

## Large File

By default, the tool manages XSSF files. If you want to manage large Excel files, you just have to call the streaming API.

```java
    PoiLight.generateStreamingExcel(outputPathExcelFile, table);
```

## Specific sheet

If you want to achieve your table on a specific spreadsheet, you can simply specify the name in the table object.

```java
    Table table = new Table();
    table.setSheetName("custom");
    PoiLight.generateExcel(outputPathExcelFile, table);
```

## Specific position

By default, your table start in A1. If you want another position, you can precise the first col and the first row of the cell at the top and left, like this :

```java
    Table table = new Table();
    table.setPosition("D8");
    PoiLight.generateExcel(outputPathExcelFile, table);
```

## Multiple table

If you want multiple tables on the same sheet, you can mix the table object and use the method createTable. This is a short example for making three tables in the same sheet. If you change the sheet name in the object table, then you can multiply table for as many spreadsheet. If you have multiple sheets, then you must close your workbook with the method Poilight.writeExcel.

```java
    Workbook wb = new XSSFWorkbook();
    Table table_1 = new Table();
    Table table_2 = new Table();
    Table table_3 = new Table();
    ....
    PoiLight.createTable(wb, table_1);
    PoiLight.createTable(wb, table_2);
    PoiLight.createTable(wb, table_3);
    PoiLight.writeExcel(wb, outputPathExcelFile);
```

## Formula

If you want to add formulas in your cells, is simple. Simply specify the boolean true when instantiating a ExcelCell.

```java
    Table table = new Table();
    table.addHeaders("ID", "NOM", "TITRE", "FORMULE");
    table.addData(new ExcelRow(1, 2, 3, new ExcelCell("SUM(D5:F5)", true)));
    table.addData(new ExcelRow(2, 10, 5641, new ExcelCell("SUM(D6:F6)", true)));
    table.addData(new ExcelRow(3, 20, 654, new ExcelCell("SUM(D7:F7)", true)));
    table.addData(new ExcelRow(4, 30, 43, new ExcelCell("SUM(D8:F8)", true)));
    
    PoiLight.generateExcel(TestHelper.getFullPath("TableNewStyleWorkbook.xlsx"), table);
```
