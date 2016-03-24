package com.github.lynchmaniac.poilight;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.github.lynchmaniac.poilight.PoiLight;
import com.github.lynchmaniac.poilight.entite.CellContent;
import com.github.lynchmaniac.poilight.entite.RowContent;
import com.github.lynchmaniac.poilight.entite.Table;
import com.github.lynchmaniac.poilight.entite.TableStyle;
import com.github.lynchmaniac.poilight.enumeration.BoardStyles;
import com.github.lynchmaniac.poilight.helpers.CreateExcelStyleHelper;

public class TestHelper {


//	private static HashMap<Integer, HashMap<Integer, Cell>> allCells = new HashMap<Integer, HashMap<Integer,Cell>>();


//	private static void getAllCells(XSSFSheet sheet, Integer firstRow, Integer firstCol) {
//		int limitRow = firstRow + 6;
//		int limitCol = firstCol + 3;
//		for (int i = firstRow; i < limitRow; i++) {
//			HashMap<Integer, Cell> colMap = new HashMap<Integer, Cell>();
//			Row currentRow = sheet.getRow(i);
//			for (int j = firstCol; j < limitCol; j++) {
//				colMap.put(j, currentRow.getCell(j));
//			}
//			allCells.put(i, colMap);
//		}
//	}

//	private static Cell getCell(Integer row, Integer col) {
//		return allCells.get(row).get(col);
//	}

//	private static void testContent(XSSFSheet sheet, Integer firstRow, Integer firstCol) {
//
//		String[] id = new String[] {"ID", "1", "2", "3", "4", "5"};
//		String[] name = new String[] {"NOM", "Henri Loevenbruck", "Cyril Massarotto", "Bernard Werber", "Maxime Chattam", "Franck Thilliez"};
//		String[] title = new String[] {"TITRE", "L'apothicaire", "Dieu est un pote à moi", "Les fourmis", "In Tenebris", "Pandemia"};
//
//		int idCol = firstCol;
//		int limit = firstRow + 6;
//		int idxExpected = 0;
//		for (int i = firstRow; i < limit; i++) {
//
//			switch (getCell(i,idCol).getCellType()) {
//			case Cell.CELL_TYPE_STRING:
//				assertEquals("Contrôle du rang " + i + " et de la colonne " + idCol, id[idxExpected], getCell(i,idCol).getStringCellValue());
//				break;
//			case Cell.CELL_TYPE_NUMERIC:
//				Double doubles = getCell(i,idCol).getNumericCellValue();
//				assertEquals("Contrôle du rang " + i + " et de la colonne " + idCol, id[idxExpected], String.valueOf(doubles.intValue()));
//				break;
//			default:
//				System.out.println();
//			}
//			idCol++;
//			assertEquals("Contrôle du rang " + i + " et de la colonne " + idCol, name[idxExpected], getCell(i,idCol).getStringCellValue());
//			idCol++;
//			assertEquals("Contrôle du rang " + i + " et de la colonne " + idCol, title[idxExpected], getCell(i,idCol).getStringCellValue());
//			idCol = firstCol;
//			idxExpected++;
//		}
//
//	}

	public static void testTable(String excelFileName, Table table) {
		testTable(excelFileName, table, false);
	}

	public static void testTable(String excelFileName, Table table, boolean isExtractStyle) {

		HashMap<BoardStyles, HashMap<String, TableStyle>>  styles = CreateExcelStyleHelper.getExcelStyle();
		try {
			XSSFWorkbook wb = new XSSFWorkbook(excelFileName);
			String sheetName = PoiLight.DEFAULT_SHEET_NAME;
			if (table.getSheetName() != null && !"".equals(table.getSheetName())) {
				sheetName = table.getSheetName();
			}
			Sheet sheet = wb.getSheet(sheetName);

			testHeader(table, isExtractStyle, styles, wb, sheet);

			// TEST OF THE BODY DATA AND STYLE
			testBody(table, isExtractStyle, styles, wb, sheet);
			wb.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static  List<String> getDataExpected() {
		
		List<String> dataExpected = new ArrayList<String>();
		dataExpected.add("1");
		dataExpected.add("Henri Loevenbruck");
		dataExpected.add("L'apothicaire");
		dataExpected.add("2");
		dataExpected.add("Cyril Massarotto");
		dataExpected.add("Dieu est un pote à moi");
		dataExpected.add("3");
		dataExpected.add("Bernard Werber");
		dataExpected.add("Les fourmis");
		dataExpected.add("4");
		dataExpected.add("Maxime Chattam");
		dataExpected.add("In Tenebris");
		dataExpected.add("5");
		dataExpected.add("Franck Thilliez");
		dataExpected.add("Pandemia");
		return dataExpected;
	}
	
	private static void testHeader(Table table, boolean isExtractStyle, HashMap<BoardStyles, HashMap<String, TableStyle>> styles, XSSFWorkbook wb, Sheet sheet) {
		String[] header = new String[] {"ID", "NOM", "TITRE"};
		Row headers = sheet.getRow(table.getRow());
		int tableIdx = table.getCol();
		for (int cellIdx = 0; cellIdx < 3; cellIdx++) {
			Cell cell = headers.getCell(tableIdx);
			CellStyle style = cell.getCellStyle();

			TableStyle headerExpectedStyle = styles.get(table.getStyle()).get("HEAD");

			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:
				Double doubles = cell.getNumericCellValue();
				assertValue(header[cellIdx], String.valueOf(doubles.intValue()), isExtractStyle);
				break;
			default:
				assertValue(header[cellIdx], cell.getStringCellValue(), isExtractStyle);
			}
			assertValue(headerExpectedStyle.getCellBorderBottom().shortValue(), style.getBorderBottom(), isExtractStyle);
			assertValue(headerExpectedStyle.getCellBorderBottom().shortValue(), style.getBorderBottom(), isExtractStyle);
			assertValue(headerExpectedStyle.getCellBorderLeft().shortValue(), style.getBorderLeft(), isExtractStyle);
			assertValue(headerExpectedStyle.getCellBorderRight().shortValue(), style.getBorderRight(), isExtractStyle);
			assertValue(headerExpectedStyle.getCellBorderTop().shortValue(), style.getBorderTop(), isExtractStyle);
			assertValue(headerExpectedStyle.getAlignment().shortValue(), style.getAlignment(), isExtractStyle); 
			assertValue(headerExpectedStyle.getBorderColor().getIndex(), style.getLeftBorderColor(), isExtractStyle);
			assertValue(headerExpectedStyle.getBorderColor().getIndex(), style.getRightBorderColor(), isExtractStyle);
			assertValue(headerExpectedStyle.getBorderColor().getIndex(), style.getBottomBorderColor(), isExtractStyle);
			assertValue(headerExpectedStyle.getBorderColor().getIndex(), style.getTopBorderColor(), isExtractStyle);

			assertValue(headerExpectedStyle.getFillColor().getIndex(), style.getFillForegroundColor(), isExtractStyle);
			Font font = wb.getFontAt(style.getFontIndex());
			assertValue(headerExpectedStyle.getFontColor().getIndex(), font.getColor(), isExtractStyle);
			assertValue(headerExpectedStyle.getFontName(), font.getFontName(), isExtractStyle);
			assertValue(headerExpectedStyle.getFontSize().shortValue(), font.getFontHeightInPoints(), isExtractStyle);
			assertValue(headerExpectedStyle.isBold(), font.getBold(), isExtractStyle);
			tableIdx++;
		}
	}

	private static void testBody(Table table, boolean isExtractStyle, HashMap<BoardStyles, HashMap<String, TableStyle>> styles,
			XSSFWorkbook wb, Sheet sheet) {
		int tableIdx;
		int begin = table.getRow() + 1;
		int end = table.getRow() + 5;
		int cellIdx = 0;
		
		List<String> dataExpected = getDataExpected();
		for (int i = begin; i < end; i++) {

			Row body = sheet.getRow(i);
			int idxEven = 0;
			tableIdx = table.getCol();
			for (int j = 0; j < 3; j++) {
				Cell cell = body.getCell(tableIdx);

				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					Double doubles = cell.getNumericCellValue();
					assertValue(dataExpected.get(cellIdx), String.valueOf(doubles.intValue()), isExtractStyle);
					break;
				default:
					assertValue(dataExpected.get(cellIdx), cell.getStringCellValue(), isExtractStyle);
				}


				CellStyle style = cell.getCellStyle();
				boolean isEven = idxEven % 2 == 0;
				TableStyle bodyExpectedStyle;
				if (isEven) {
					bodyExpectedStyle = styles.get(table.getStyle()).get("BODY_EVEN");
				} else {
					bodyExpectedStyle = styles.get(table.getStyle()).get("BODY_ODD");
				}

				assertValue(bodyExpectedStyle.getCellBorderBottom().shortValue(), style.getBorderBottom(), isExtractStyle);
				assertValue(bodyExpectedStyle.getCellBorderLeft().shortValue(), style.getBorderLeft(), isExtractStyle);
				assertValue(bodyExpectedStyle.getCellBorderRight().shortValue(), style.getBorderRight(), isExtractStyle);
				assertValue(bodyExpectedStyle.getCellBorderTop().shortValue(), style.getBorderTop(), isExtractStyle);
				assertValue(bodyExpectedStyle.getAlignment().shortValue(), style.getAlignment(), isExtractStyle); 
				assertValue(bodyExpectedStyle.getBorderColor().getIndex(), style.getLeftBorderColor(), isExtractStyle);
				assertValue(bodyExpectedStyle.getBorderColor().getIndex(), style.getRightBorderColor(), isExtractStyle);
				assertValue(bodyExpectedStyle.getBorderColor().getIndex(), style.getBottomBorderColor(), isExtractStyle);
				assertValue(bodyExpectedStyle.getBorderColor().getIndex(), style.getTopBorderColor(), isExtractStyle);

				assertValue(bodyExpectedStyle.getFillColor().getIndex(), style.getFillForegroundColor(), isExtractStyle);
				Font font = wb.getFontAt(style.getFontIndex());
				assertValue(bodyExpectedStyle.getFontColor().getIndex(), font.getColor(), isExtractStyle);
				assertValue(bodyExpectedStyle.getFontName(), font.getFontName(), isExtractStyle);
				assertValue(bodyExpectedStyle.getFontSize().shortValue(), font.getFontHeightInPoints(), isExtractStyle);
				assertValue(bodyExpectedStyle.isBold(), font.getBold(), isExtractStyle);
				cellIdx++;
				tableIdx++;
			}

		}
	}

	private static void assertValue(Object expected, Object value, boolean isExtractStyle) {
		if (isExtractStyle) {
			System.out.println(value);
		} else {
			assertEquals(expected, value);
		}
	}




//	private static XSSFWorkbook getTestWorkbook(String excelFilename) {
//		XSSFWorkbook wb = null;
//		try {
//			wb = new XSSFWorkbook(excelFilename);
//
//		} catch (IOException e) {
//			assertFalse(e.getMessage(), true);
//		}
//		return wb;
//	}
//	private static XSSFSheet getSheet(XSSFWorkbook wb, String sheetName) {
//		//Get first sheet from the workbook
//		return wb.getSheet(sheetName);
//	}

//	private static void testBoardData(XSSFSheet sheet, Integer firstRow, Integer firstCol) {
//		allCells = new HashMap<Integer, HashMap<Integer,Cell>>();
//		firstRow = firstRow - 1;
//		firstCol = firstCol - 1;
//		getAllCells(sheet, firstRow, firstCol);
//		testContent(sheet, firstRow, firstCol);
//	}

//	static void controlResultTest(String excelFilename) {
//		controlResultTest(excelFilename, PoiLight.DEFAULT_SHEET_NAME, 1, 1);
//	}


//	private static void controlResultTest(String excelFilename, String sheetName, Integer firstRow, Integer firstCol) {
//		XSSFWorkbook wb = getTestWorkbook(excelFilename);
//		XSSFSheet sheet = getSheet(wb, sheetName);
//		testBoardData(sheet, firstRow, firstCol);
//	}

//	static void controlResultTest(String excelFilename, String sheetName, List<Coordinates> coords) {
//		XSSFWorkbook wb = getTestWorkbook(excelFilename);
//		XSSFSheet sheet = getSheet(wb, sheetName);
//		for (Coordinates coord : coords) {
//			testBoardData(sheet, coord.getRow(), coord.getCol());
//		}
//	}


	static Table getTable(String sheetName, BoardStyles bs, String position) {
		Table table = getTable();
		table.setSheetName(sheetName);
		table.setStyle(bs);
		table.setPosition(position);
		return table;
	}

	static Table getTable() {
		Table table = new Table();
		table.addHeader(new CellContent("ID"));
		table.addHeader(new CellContent("NOM"));
		table.addHeader(new CellContent("TITRE"));
		table.addData(new RowContent(new CellContent(1), new CellContent("Henri Loevenbruck"), new CellContent("L'apothicaire")));
		table.addData(new RowContent(new CellContent(2), new CellContent("Cyril Massarotto"), new CellContent("Dieu est un pote à moi")));
		table.addData(new RowContent(new CellContent(3), new CellContent("Bernard Werber"), new CellContent("Les fourmis")));
		table.addData(new RowContent(new CellContent(4), new CellContent("Maxime Chattam"), new CellContent("In Tenebris")));
		table.addData(new RowContent(new CellContent(5), new CellContent("Franck Thilliez"), new CellContent("Pandemia")));

		return table;
	}

}
