# POILIGHT
[![Build Status](https://travis-ci.org/lynchmaniac/poilight.svg?branch=master)](https://travis-ci.org/lynchmaniac/poilight)

Poilight is a wrapper of POI to accellerate the generation of Excel's files.

## Generate a simple File
The principle of poilight is to simply manage Excel spreadsheets. It's simple tables corresponding to the output of your Java applications. You must set a complete full path for your file and fill some data. For that, you can use the Table object, see the next chapter.

```java
		String outputPathExcelFile = "file.xlsx";
		Table table = new Table();
		PoiLight.generateExcel(outputPathExcelFile, table);
```
## Structure of the data
The data is a Table object. It is the table you want to achieve. If you want to add a headers, you can precise them with the addHeader method. It take a ExcelCell. ExcelCell is a structure in which you can precise the value and the of the cell. We will see this in a later chapter.
You can pass data to the table object with the method addData which take a ExcelRow as parameter. ExcelRow is just a list of ExcelCell.
This an example of a basic table
```java
		Table table = new Table();
		table.addHeader(new ExcelCell("ID"));
		table.addHeader(new ExcelCell("NOM"));
		table.addHeader(new ExcelCell("TITRE"));
		table.addData(new ExcelRow(new ExcelCell(1), new ExcelCell("Henri Loevenbruck"), new ExcelCell("L'apothicaire")));
		table.addData(new ExcelRow(new ExcelCell(2), new ExcelCell("Cyril Massarotto"), new ExcelCell("Dieu est un pote à moi")));
		table.addData(new ExcelRow(new ExcelCell(3), new ExcelCell("Bernard Werber"), new ExcelCell("Les fourmis")));
		table.addData(new ExcelRow(new ExcelCell(4), new ExcelCell("Maxime Chattam"), new ExcelCell("In Tenebris")));
		table.addData(new ExcelRow(new ExcelCell(5), new ExcelCell("Franck Thilliez"), new ExcelCell("Pandemia")));
```

## Installation
Poilight is an artefact Maven so you can put this in your pom.xml to have Poilight in your project.
```java
  	<dependencies>
		<dependency>
			<groupId>com.github.lynchmaniac</groupId>
			<artifactId>poilight</artifactId>
			<version>0.1</version>
		</dependency>
	</dependencies>
```
## Large File
By default, the tool manages XSSF files. If you want to manage large Excel files, you just have to call the streaming API.
```java
		PoiLight.generateStreamingExcel(outputPathExcelFile, table);
```

## The predefined styles
Poilight embeds the entire 60 preset styles in Excel. You can specifie a style with the enum BoardStyles.
```java
		Table table = new Table();
		table.setStyle(BoardStyles.BOARD_LIGHT_RED_3_STYLE);
		PoiLight.generateExcel(outputPathExcelFile, table);
```
You can find in the example directory a Excel File with the 60 styles.

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
If you want multiple table on the same sheet, you can combine the table object and use the method createTable. This is a short example for making three tables in the same sheet. If you change the sheet name in the object table, then you can multiple table on multiple spreadsheet. If you have multiple sheet, then you must close your workbook with the metho Poilight.writeExcel.
```java
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, outputPathExcelFile);
```
## Specific style

If you don't want a predifined style, you can put a style on each cell by using the CellContent with a CellStyle from POI, like this

<pre><code>
		Table table = new Table();
		table.addHeader(new ExcelCell("ID"));
		table.addHeader(new ExcelCell("NOM"));
		table.addHeader(new ExcelCell("TITRE"));
		<b>CellStyle cs = wb.createCellStyle();</b>
		<b>cs.setFillForegroundColor(StyleHelper.getColor(128, 100, 162).getIndex());</b>
		<b>cs.setFillPattern(CellStyle.SOLID_FOREGROUND);</b>

		table.addData(new ExcelRow(new ExcelCell(4<b>, cs</b>), new ExcelCell("Maxime Chattam"), new ExcelCell("In Tenebris")));
		table.addData(new ExcelRow(new ExcelCell(5), new ExcelCell("Franck Thilliez"), new ExcelCell("Pandemia")));
</code></pre>
Here we have a specific style on the first cell. The background is set to purple. With this mechanism you can customize style as you like.
