# POILIGHT
Poilight is a wrapper of POI to accellerate the generation of Excel's files.

## Generate a simple File
The principle of poilight is to simply manage Excel spreadsheets. It is simple tables corresponding to the output of your Java applications.
For this it is necessary to pass your data as a HashMap where each line corresponds to a row of your table.
```java
		String excelPathFile = "d:\\PoiLightFile.xlsx";
		PoiLight.generateExcel(excelPathFile, table);
```
## Structure of the data
The data is a Table object. It is the table you want to achieve. If you want to add a headers, you can precise them with the addHeader method. It take a CellContent. CellContent is a structure in which you can precise the value of the cell. You can also precise a style for this specific cell. We will see this in a later chapter.
You can pass data to the table object with the method addData which take a RowContent as parameter. RowContent is just a list of CellContent.
This an example of a basic table
```java
		Table table = new Table();
		table.addHeader(new CellContent("ID"));
		table.addHeader(new CellContent("NOM"));
		table.addHeader(new CellContent("TITRE"));
		table.addData(new RowContent(new CellContent(1), new CellContent("Henri Loevenbruck"), new CellContent("L'apothicaire")));
		table.addData(new RowContent(new CellContent(2), new CellContent("Cyril Massarotto"), new CellContent("Dieu est un pote à moi")));
		table.addData(new RowContent(new CellContent(3), new CellContent("Bernard Werber"), new CellContent("Les fourmis")));
		table.addData(new RowContent(new CellContent(4), new CellContent("Maxime Chattam"), new CellContent("In Tenebris")));
		table.addData(new RowContent(new CellContent(5), new CellContent("Franck Thilliez"), new CellContent("Pandemia")));
```

## Large File
By default, the tool manages XSSF files. If you want to manage large Excel files, you just have to call the streaming API.
```java
		String excelPathFile = "d:\\PoiLightBigFile.xlsx";
		PoiLight.generateStreamingExcel(excelPathFile, table);
```

## The predefined styles
Poilight embeds the entire 60 preset styles in Excel. You can specifie a style with the enum BoardStyles.
```java
		Table table = new Table();
		table.setStyle(BoardStyles.BOARD_LIGHT_RED_3_STYLE);
		PoiLight.generateExcel(excelPathFile, table);
```
You can find in the example directory a Excel File with the 60 styles.

## Specific sheet
If you want to achieve your table on a specific spreadsheet, you can simply specify the name in the table object.
```java
		Table table = new Table();
		table.setSheetName("custom");
		PoiLight.generateExcel(excelPathFile, table);
```

## Specific poistion
By default, your table start in A1.
