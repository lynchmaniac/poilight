# poilight
Poilight is a wrapper of POI to accellerate the generation of Excel's files.

## Generate a simple File
The principle of poilight is to simply manage Excel spreadsheets. It is simple tables corresponding to the output of your Java applications.
For this it is necessary to pass your data as a HashMap where each line corresponds to a row of your table.
```java
		String excelPathFile = "d:\\PoiLightFile.xlsx";
		PoiLight.generateExcel(excelPathFile, getData());
```
## Structure of the data
The data must be a LinkedHashMap<Integer, RowContent>. The Integer is the number in the row of the table. 
RowContent is a structure for storing data and the style of each cell.
```java
		RowContent rowContent = new RowContent();
		rowContent.addValue(new CellContent("AUTHOR"));
		rowContent.addValue(new CellContent("TITLE"));
		hashMap.put(0, rowContent);
		
		rowContent = new RowContent(); 
		rowContent.addValue(new CellContent("Cyril Massarotto"));
		rowContent.addValue(new CellContent("Dieu est un pote à moi"));
		hashMap.put(1, rowContent);
  		
		rowContent = new RowContent(); 
		rowContent.addValue(new CellContent("Henri Loevenbruck"));
		rowContent.addValue(new CellContent("L'apothicaire"));
		hashMap.put(2, rowContent);
```

## Large File
By default, the tool manages XSSF files. If you want to manage large Excel files, you just have to call the streaming API.
```java
		String excelPathFile = "d:\\PoiLightBigFile.xlsx";
		PoiLight.generateStreamingExcel(excelPathFile, getData());
```
