package com.github.lynchmaniac.poilight;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

import java.io.File;
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
import com.github.lynchmaniac.poilight.enumerations.BoardStyles;
import com.github.lynchmaniac.poilight.helpers.CreateExcelStyleHelper;
import com.github.lynchmaniac.poilight.models.ExcelCell;
import com.github.lynchmaniac.poilight.models.ExcelRow;
import com.github.lynchmaniac.poilight.models.Table;
import com.github.lynchmaniac.poilight.models.TableStyle;

public class TestHelper {

	protected static String getFullPath(String fileName) {
		return "src" + File.separator + "test" + File.separator + "resources" + File.separator + fileName;
	}

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
			assertFalse(true);
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

	static Table getTable(String sheetName, BoardStyles bs, String position) {
		Table table = getTable();
		table.setSheetName(sheetName);
		table.setStyle(bs);
		table.setPosition(position);
		return table;
	}

	static Table getTable() {
		Table table = new Table();
		table.addHeader(new ExcelCell("ID"));
		table.addHeader(new ExcelCell("NOM"));
		table.addHeader(new ExcelCell("TITRE"));
		table.addData(new ExcelRow(new ExcelCell(1), new ExcelCell("Henri Loevenbruck"), new ExcelCell("L'apothicaire")));
		table.addData(new ExcelRow(new ExcelCell(2), new ExcelCell("Cyril Massarotto"), new ExcelCell("Dieu est un pote à moi")));
		table.addData(new ExcelRow(new ExcelCell(3), new ExcelCell("Bernard Werber"), new ExcelCell("Les fourmis")));
		table.addData(new ExcelRow(new ExcelCell(4), new ExcelCell("Maxime Chattam"), new ExcelCell("In Tenebris")));
		table.addData(new ExcelRow(new ExcelCell(5), new ExcelCell("Franck Thilliez"), new ExcelCell("Pandemia")));

		return table;
	}

}
