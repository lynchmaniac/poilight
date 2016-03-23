/**
 * 
 */
package fr.lynchmaniac.tools.poilight;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import fr.lynchmaniac.tools.poilight.enumeration.BoardStyles;

/**
 * @author piard
 * @since 0.1
 *
 */
public class DarkStyleTest {
	
	
	@Test
	public void darkGrayWorkbook() {
		String excelFilename = "d:\\tmp\\DarkGrayWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "A1"));
	}
	
	@Test
	public void darkBlueWorkbook() {
		String excelFilename = "d:\\tmp\\DarkBlueWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "A1"));
		
	}
	
	
	@Test
	public void darkRedWorkbook() {
		String excelFilename = "d:\\tmp\\DarkRedWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_RED_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_RED_1_STYLE, "A1"));
	}
	
	@Test
	public void darkGreenWorkbook() {
		String excelFilename = "d:\\tmp\\DarkGreenWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "A1"));
	}
	
	@Test
	public void darkPurpleWorkbook() {
		String excelFilename = "d:\\tmp\\DarkPurpleWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_PURPLE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_PURPLE_1_STYLE, "A1"));
	}
	
	@Test
	public void darkTurquoiseWorkbook() {
		String excelFilename = "d:\\tmp\\DarkTurquoiseWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, "A1"));
	}
	
	@Test
	public void darkOrangeWorkbook() {
		String excelFilename = "d:\\tmp\\DarkOrangeWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_ORANGE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_ORANGE_1_STYLE, "A1"));
	}
	
	
	@Test
	public void darkMix1Workbook() {
		String excelFilename = "d:\\tmp\\DarkMix1Workbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_1_STYLE, "A1"));
	}
	
	@Test
	public void darkMix2Workbook() {
		String excelFilename = "d:\\tmp\\DarkMix2Workbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_2_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_2_STYLE, "A1"));
	}
	
	@Test
	public void darkMix3Workbook() {
		String excelFilename = "d:\\tmp\\DarkMix3Workbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_3_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_3_STYLE, "A1"));
	}
	
	@Test
	public void darkMix4Workbook() {
		String excelFilename = "d:\\tmp\\DarkMix4Workbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_4_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_4_STYLE, "A1"));
	}
	
	
	@Test
	public void darkWorkbook() {
		String excelFilename = "d:\\tmp\\DarkWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();

		
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "B2"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "F2"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_RED_1_STYLE, "J2"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "N2"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "R2"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "V2"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "Z2"));
		
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_1_STYLE, "B15"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_2_STYLE, "F15"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_3_STYLE, "J15"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_4_STYLE, "N15"));
		
		PoiLight.writeExcel(wb, excelFilename);
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "B2"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "F2"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_RED_1_STYLE, "J2"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "N2"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "R2"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "V2"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "Z2"));
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_1_STYLE, "B15"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_2_STYLE, "F15"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_3_STYLE, "J15"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_4_STYLE, "N15"));
	}
}
