package fr.lynchmaniac.tools.poilight;

import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import fr.lynchmaniac.tools.poilight.entite.CellContent;
import fr.lynchmaniac.tools.poilight.entite.RowContent;
import fr.lynchmaniac.tools.poilight.entite.Table;
import fr.lynchmaniac.tools.poilight.enumeration.BoardStyles;

/**
 * This class is the entry point of the application. It contains methods to generate 
 * Excel spreadsheets as well as methods to generate complete Excel files.
 * If the user does not specify a style in leaf cells then the default style is applied. 
 * Otherwise it can use the predefined styles from Excel 2016. If it specifies a 
 * style on cells, it has priority over the predefined styles.
 * 
 * @author vpiard
 * @since 0.1
 */
public class PoiLight {

	public static final String DEFAULT_SHEET_NAME = "data";

	/**
	 * Generate an Excel file.
	 * 
	 * @param filePath the full path where the file should be saved
	 * @param data the data to be written in the spreadsheet
	 */
	public static void generateExcel(String filePath, Table data){
		XSSFWorkbook wb = new XSSFWorkbook();
		createTable(wb, data, false);
		writeExcel(wb, filePath);
	}
	
//	/**
//	 * Generate an Excel file.
//	 * 
//	 * @param filePath the full path where the file should be saved
//	 * @param data the data to be written in the spreadsheet
//	 */
//	public static void generateExcel(String filePath, LinkedHashMap<Integer, RowContent> data){
//		XSSFWorkbook wb = new XSSFWorkbook();
//		generateExcel(wb, filePath, data, false);
//	}
	
//	/**
//	 * Generate an Excel file.
//	 * 
//	 * @param filePath the full path where the file should be saved
//	 * @param data the data to be written in the spreadsheet
//	 * @param bs all properties suitable on the style of a cell
//	 */
//	public static void generateExcel(String filePath, LinkedHashMap<Integer, RowContent> data, BoardStyles bs){
//		XSSFWorkbook wb = new XSSFWorkbook();
//		generateExcel(wb, filePath, data, false, null, bs);
//	}
	
//	/**
//	 * Generate an Excel file.
//	 * 
//	 * @param filePath the full path where the file should be saved
//	 * @param data the data to be written in the spreadsheet
//	 * @param bs all properties suitable on the style of a cell
//	 * @param firstRow the leading index of the table including the headers
//	 * @param firstCol the index of the first column
//	 */
//	public static void generateExcel(String filePath, LinkedHashMap<Integer, RowContent> data, BoardStyles bs, int firstRow, int firstCol){
//		XSSFWorkbook wb = new XSSFWorkbook();
//		generateExcel(wb, filePath, data, false, null, bs, firstRow, firstCol);
//	}
	
	/**
	 * Generate an Excel file in streaming mode. Use for large file.
	 * 
	 * @param filePath the full path where the file should be saved
	 * @param data the data to be written in the spreadsheet
	 */
	public static void generateStreamingExcel(String filePath, Table data){
		SXSSFWorkbook wb = new SXSSFWorkbook();
		createTable(wb, data, true);
		writeExcel(wb, filePath);
	}
//	public static void generateStreamingExcel(String filePath, LinkedHashMap<Integer, RowContent> data){
//		SXSSFWorkbook wb = new SXSSFWorkbook();
//		generateExcel(wb, filePath, data, true);
//	}
	
//	/**
//	 * Generate an Excel file in streaming mode. Use for large file.
//	 * 
//	 * @param filePath the full path where the file should be saved
//	 * @param data the data to be written in the spreadsheet
//	 * @param bs all properties suitable on the style of a cell
//	 */
//	public static void generateStreamingExcel(String filePath, LinkedHashMap<Integer, RowContent> data, BoardStyles bs){
//		SXSSFWorkbook wb = new SXSSFWorkbook();
//		generateExcel(wb, filePath, data, true, null, bs);
//	}
	
	
//	/**
//	 * Generate an Excel file.
//	 * 
//	 * @param wb Excel Workbook
//	 * @param filePath the full path where the file should be saved
//	 * @param data the data to be written in the spreadsheet
//	 * @param isStreaming indicates whether to write the streaming file
//	 */
//	public static void generateExcel(Workbook wb, String filePath, LinkedHashMap<Integer, RowContent> data, boolean isStreaming){
//		generateExcel(wb, filePath, data, isStreaming, null, null);
//	}

//	/**
//	 * Generate an Excel file.
//	 * 
//	 * @param wb Excel Workbook
//	 * @param filePath the full path where the file should be saved
//	 * @param data the data to be written in the spreadsheet
//	 * @param isStreaming indicates whether to write the streaming file
//	 * @param sheetName the name of the spreadsheet
//	 * @param bs all properties suitable on the style of a cell
//	 */
//	public static void generateExcel(Workbook wb, String filePath, LinkedHashMap<Integer, RowContent> data, boolean isStreaming, String sheetName, BoardStyles bs) {
//		generateExcel(wb, filePath, data, isStreaming, sheetName, bs, 1, 1);
//	}
	
//	/**
//	 * Generate an Excel file.
//	 * 
//	 * @param wb Excel Workbook
//	 * @param filePath the full path where the file should be saved
//	 * @param data the data to be written in the spreadsheet
//	 * @param isStreaming indicates whether to write the streaming file
//	 * @param sheetName the name of the spreadsheet
//	 * @param bs all properties suitable on the style of a cell
//	 * @param firstRow the leading index of the table including the headers
//	 * @param firstCol the index of the first column
//	 */
//	public static void generateExcel(Workbook wb, String filePath, LinkedHashMap<Integer, RowContent> data, boolean isStreaming, String sheetName, BoardStyles bs, int firstRow, int firstCol) {
//
//		createTable(wb, data, isStreaming, null, bs, firstRow, firstCol);
//		writeExcel(wb, filePath);
//	}
	
	/**
	 * Write a workbook in a file.
	 * 
	 * @param wb Excel Workbook
	 * @param filePath the full path where the file should be saved
	 */
	public static void writeExcel(Workbook wb, String filePath) {
		try {
			FileOutputStream fileOut = new FileOutputStream(filePath);

			wb.write(fileOut);
			fileOut.close();
		}
		catch (Exception e) {
			System.out.println(e);
		}
	}
	

//	/**
//	 * Generate an Excel spreadsheet.
//	 * 
//	 * @param wb Excel Workbook
//	 * @param data the data to be written in the spreadsheet
//	 * @param isStreaming indicates whether to write the streaming file
//	 */
//	public static void createTable(Workbook wb, LinkedHashMap<Integer, RowContent> data, boolean isStreaming) {
//		createTable(wb, data, isStreaming, null, null);
//	}

//	/**
//	 * Generate an Excel spreadsheet.
//	 * 
//	 * @param wb Excel Workbook
//	 * @param data the data to be written in the spreadsheet
//	 * @param isStreaming indicates whether to write the streaming file
//	 * @param sheetName the name of the spreadsheet
//	 * @param bs all properties suitable on the style of a cell
//	 */
//	public static void createTable(Workbook wb, LinkedHashMap<Integer, RowContent> data, boolean isStreaming, String sheetName, BoardStyles bs) {
//		createTable(wb, data, isStreaming, sheetName, bs, 0, 0);	
//	}
	
//	/**
//	 * Generate an Excel spreadsheet.
//	 * 
//	 * @param wb Excel Workbook
//	 * @param data the data to be written in the spreadsheet
//	 * @param isStreaming indicates whether to write the streaming file
//	 * @param sheetName the name of the spreadsheet
//	 * @param bs all properties suitable on the style of a cell
//	 * @param firstRow the leading index of the table including the headers
//	 * @param firstCol the index of the first column
//	 */
//	public static void createTable(Workbook wb, LinkedHashMap<Integer, RowContent> data, boolean isStreaming, String sheetName, BoardStyles bs, Integer firstRow, Integer firstCol) {
//		//New Sheet
//		if (sheetName == null || "".equals(sheetName)) {
//			sheetName = DEFAULT_SHEET_NAME;
//		}
//		
//		Sheet sheet = wb.getSheet(sheetName);
//		if (sheet == null) {
//			sheet = isStreaming? (SXSSFSheet) wb.createSheet(sheetName) : (XSSFSheet) wb.createSheet(sheetName);
//		}
//
//		// Control of position index
//		firstRow = controlRowIndex(wb, data, firstRow);
//		firstCol = controlColIndex(wb, data, firstCol);
//		
//		// Creation of style
//		if (bs == null) {
//			bs = BoardStyles.BOARD_DEFAULT_STYLE;
//		}
//		// Row and column indexes
//		int idx = firstRow - 1;
//		int total = data.size();
//		int idxEven = 1;
//		// Generate column headings
//		for (Entry<Integer, RowContent> entry : data.entrySet()) {
//			Integer key = entry.getKey();
//			RowContent currentRowContent = data.get(key); 
//			
//			Row row = sheet.getRow(idx);
//			if (row == null) {
//				row = sheet.createRow(idx);
//			}
//			int i = firstCol -1;
//			for (CellContent cell : currentRowContent.getValue()) {
//				Cell c = row.getCell(i);
//				if (c == null) {
//					c = row.createCell(i);
//				}
//				setCellValue(c, cell);
//				boolean isHeader = idx == (firstRow - 1);
//				boolean isFooter = total == idxEven;
//				boolean isEven = !(idxEven % 2 == 0);
//				applyCellStyle(wb, cell, c, bs, isHeader, isFooter, isEven);
//				i++;
//			} 
//			idxEven++;
//			idx++;
//		}
//	}
	
	
	public static void createTable(Workbook wb, Table data) {
		createTable(wb, data, false);
	}
	public static void createStreamingTable(Workbook wb, Table data) {
		createTable(wb, data, true);
	}
	/**
	 * Generate an Excel spreadsheet.
	 * 
	 * @param wb Excel Workbook
	 * @param data the data to be written in the spreadsheet
	 * @param isStreaming indicates whether to write the streaming file
	 * @param sheetName the name of the spreadsheet
	 * @param bs all properties suitable on the style of a cell
	 * @param firstRow the leading index of the table including the headers
	 * @param firstCol the index of the first column
	 */
	private static void createTable(Workbook wb, Table data, boolean isStreaming) {
		//New Sheet
		if (data.getSheetName() == null || "".equals(data.getSheetName())) {
			data.setSheetName(DEFAULT_SHEET_NAME);
		}
		
		Sheet sheet = wb.getSheet(data.getSheetName());
		if (sheet == null) {
			sheet = isStreaming? (SXSSFSheet) wb.createSheet(data.getSheetName()) : (XSSFSheet) wb.createSheet(data.getSheetName());
		}

		// Control of position index
		data.setRow(controlRowIndex(wb, data.getData(), data.getRow()));
		data.setCol(controlColIndex(wb, data.getData(), data.getCol()));
		
		// Creation of style
		if (data.getStyle() == null) {
			data.setStyle(BoardStyles.BOARD_DEFAULT_STYLE);
		}
		// Row and column indexes
		int idx = data.getRow();
		//int idx = data.getRow() - 1;
		int total = data.getData().size();
		int idxEven = 1;
		
		// Generate header
		if (data.hasHeader()) {
			Row row = sheet.getRow(idx);
			if (row == null) {
				row = sheet.createRow(idx);
			}
			int i = data.getCol();
			for (CellContent cell : data.getHeader()) {
				Cell c = row.getCell(i);
				if (c == null) {
					c = row.createCell(i);
				}
				setCellValue(c, cell);
				applyCellStyle(wb, cell, c, data.getStyle(), true, false, true);
				i++;
			}
			idx++;
		}
		
		// Generate column headings
		for (RowContent currentRowContent : data.getData()) {
			
			Row row = sheet.getRow(idx);
			if (row == null) {
				row = sheet.createRow(idx);
			}
			int i = data.getCol();
			for (CellContent cell : currentRowContent.getValue()) {
				Cell c = row.getCell(i);
				if (c == null) {
					c = row.createCell(i);
				}
				setCellValue(c, cell);
//				boolean isHeader = idx == (data.getRow() - 1);
				boolean isFooter = total == idxEven;
				boolean isEven = !(idxEven % 2 == 0);
				applyCellStyle(wb, cell, c, data.getStyle(), false, isFooter, isEven);
				i++;
			} 
			idxEven++;
			idx++;
		}
	}
	
//	private void createRow(Sheet sheet, int idx) {
//		Row row = sheet.getRow(idx);
//		if (row == null) {
//			row = sheet.createRow(idx);
//		}
//		int i = firstCol -1;
//		for (CellContent cell : currentRowContent.getValue()) {
//			Cell c = row.getCell(i);
//			if (c == null) {
//				c = row.createCell(i);
//			}
//			setCellValue(c, cell);
//			boolean isHeader = idx == (firstRow - 1);
//			boolean isFooter = total == idxEven;
//			boolean isEven = !(idxEven % 2 == 0);
//			applyCellStyle(wb, cell, c, bs, isHeader, isFooter, isEven);
//			i++;
//		} 
//		idxEven++;
//		idx++;
//	}

	/**
	 * check the index of the first row. If the index is not included in the possible limits of Excel, 
	 * it is then corrected to match the closest.
	 * HSSF a type of workbook can only contain 65,536 rows.
	 * XSSF a type of workbook or SXSSF only contain rows 1048576.
	 * If the index is less than 1 then it is reduced to 1. If it is greater than the maximum terminal 
	 * then returns the max index along the length of the table.
	 * 
	 * @param wb Excel Workbook
	 * @param data the data to be written in the spreadsheet
	 * @param firstRow the leading index of the table including the headers
	 * @return controlled the index
	 */
//	public static Integer controlRowIndex(Workbook wb, LinkedHashMap<Integer, RowContent> data, Integer firstRow) {
//		Integer result = firstRow;
//		Integer limit = firstRow + data.size();
//		if (firstRow < 1) {
//			result = 1;
//		} else if (wb instanceof HSSFWorkbook && limit >= PoiLightConst.EXCEL_1997_MAX_ROW) {
//			result = PoiLightConst.EXCEL_1997_MAX_ROW - data.size();
//		} else if ((wb instanceof XSSFWorkbook || wb instanceof SXSSFWorkbook) && limit >= PoiLightConst.EXCEL_2007_MAX_ROW) {
//			result = PoiLightConst.EXCEL_2007_MAX_ROW - data.size();
//		}
//		return result;
//	}
	private static Integer controlRowIndex(Workbook wb, List<RowContent> data, Integer firstRow) {
		Integer result = firstRow;
		Integer limit = firstRow + data.size();
		if (firstRow < 0) {
			result = 0;
		} else if (wb instanceof HSSFWorkbook && limit >= PoiLightConst.EXCEL_1997_MAX_ROW) {
			result = PoiLightConst.EXCEL_1997_MAX_ROW - data.size();
		} else if ((wb instanceof XSSFWorkbook || wb instanceof SXSSFWorkbook) && limit >= PoiLightConst.EXCEL_2007_MAX_ROW) {
			result = PoiLightConst.EXCEL_2007_MAX_ROW - data.size();
		}
		return result;
	}
	
	/**
	 * check the index of the first column. If the index is not included in the possible limits of Excel,
	 *  it is then corrected to match the closest.
	 *  HSSF a type of workbook can contain only 256 columns.
	 *  XSSF a type of workbook or SXSSF contain only 16385 columns.
	 *  If the index is less than 1 then it is reduced to 1. If it is greater than the maximum terminal 
	 *  then returns the index max following the width of the table.
	 *  
	 * @param wb Excel Workbook
	 * @param data the data to be written in the spreadsheet
	 * @param firstCol the index of the first column
	 * @return controlled the index
	 */
//	private static Integer controlColIndex(Workbook wb, LinkedHashMap<Integer, RowContent> data, Integer firstCol) {
//		Integer result = firstCol;
//		Integer limit = firstCol + data.get(0).getValue().size();
//		if (firstCol < 1) {
//			result = 1;
//		} else if (wb instanceof HSSFWorkbook && limit >= PoiLightConst.EXCEL_1997_MAX_COL) {
//			result = PoiLightConst.EXCEL_1997_MAX_COL - data.get(0).getValue().size();
//		} else if ((wb instanceof XSSFWorkbook || wb instanceof SXSSFWorkbook) && limit >= PoiLightConst.EXCEL_2007_MAX_COL) {
//			result = PoiLightConst.EXCEL_2007_MAX_COL - data.get(0).getValue().size();
//		}
//		return result;
//	}	
	private static Integer controlColIndex(Workbook wb, List<RowContent> data, Integer firstCol) {
		Integer result = firstCol;
		Integer limit = firstCol + data.get(0).getValue().size();
		if (firstCol < 0) {
			result = 0;
		} else if (wb instanceof HSSFWorkbook && limit >= PoiLightConst.EXCEL_1997_MAX_COL) {
			result = PoiLightConst.EXCEL_1997_MAX_COL - data.get(0).getValue().size();
		} else if ((wb instanceof XSSFWorkbook || wb instanceof SXSSFWorkbook) && limit >= PoiLightConst.EXCEL_2007_MAX_COL) {
			result = PoiLightConst.EXCEL_2007_MAX_COL - data.get(0).getValue().size();
		}
		return result;
	}	
	/**
	 * Applies the style to the cell. If the Cell Content object contains a 
	 * style then it is applied, it takes precedence over the rest. Otherwise 
	 * a predefined style is applied.
	 * 
	 * @param wb Excel Workbook
	 * @param cell the content of the cell
	 * @param c the cell
	 * @param style all properties suitable on the style of a cell
	 * @param isHeader indicates whether it is a header
	 * @param isFooter indicates whether it is a footer
	 * @param isEven indicates whether the current cell is on a par or odd row
	 */
	private static void applyCellStyle(Workbook wb, CellContent cell, Cell c, BoardStyles style, boolean isHeader, boolean isFooter, boolean isEven) {
		if (cell.getStyle() != null) {
			c.setCellStyle(cell.getStyle());
		} else {
			// on applique un style prédéfini
			if (isHeader) {
				// c'est le header le style est différent
				c.setCellStyle(PoiLightStyle.getHeaderStyle(wb, style));
			} else if(isFooter) {
				// c'est le footer le style est différent
				c.setCellStyle(PoiLightStyle.getFooterStyle(wb, style, isEven));
			}else {
				c.setCellStyle(PoiLightStyle.getBodyStyle(wb, style, isEven));
			}
		}
	}

	/**
	 * Fill the cell value.
	 * 
	 * @param c the cell
	 * @param cell the content of the cell
	 */
	private static void setCellValue(Cell c, CellContent cell) {
		if (cell.getValue() instanceof String) {
			c.setCellValue((String) cell.getValue());
		}
		if (cell.getValue() instanceof Integer) {
			c.setCellValue((Integer) cell.getValue());
		}
		if (cell.getValue() instanceof Long) {
			c.setCellValue((Long) cell.getValue());
		}
		if (cell.getValue() instanceof Double) {
			c.setCellValue((Double) cell.getValue());
		}
	}
	
}
