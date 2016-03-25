/**
 * 
 */
package com.github.lynchmaniac.poilight;

import java.io.File;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Test;

import com.github.lynchmaniac.poilight.PoiLight;
import com.github.lynchmaniac.poilight.enumerations.BoardStyles;

/**
 * @author piard
 * @since 0.1
 *
 */
public class DarkStyleTest {
	
	@BeforeClass
	public static void cleanRessources() {
		new File(TestHelper.getFullPath("DarkGrayWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkBlueWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkRedWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkGreenWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkPurpleWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkTurquoiseWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkOrangeWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkMix1Workbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkMix2Workbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkMix3Workbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkMix4Workbook.xlsx")).delete();
		new File(TestHelper.getFullPath("DarkWorkbook.xlsx")).delete();
	}
	
	@Test
	public void darkGrayWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkGrayWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "A1"));
	}
	
	@Test
	public void darkBlueWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkBlueWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "A1"));
		
	}
	
	
	@Test
	public void darkRedWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkRedWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_RED_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_RED_1_STYLE, "A1"));
	}
	
	@Test
	public void darkGreenWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkGreenWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "A1"));
	}
	
	@Test
	public void darkPurpleWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkPurpleWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_PURPLE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_PURPLE_1_STYLE, "A1"));
	}
	
	@Test
	public void darkTurquoiseWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkTurquoiseWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, "A1"));
	}
	
	@Test
	public void darkOrangeWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkOrangeWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_ORANGE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_ORANGE_1_STYLE, "A1"));
	}
	
	
	@Test
	public void darkMix1Workbook() {
		String excelFilename = TestHelper.getFullPath("DarkMix1Workbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_1_STYLE, "A1"));
	}
	
	@Test
	public void darkMix2Workbook() {
		String excelFilename = TestHelper.getFullPath("DarkMix2Workbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_2_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_2_STYLE, "A1"));
	}
	
	@Test
	public void darkMix3Workbook() {
		String excelFilename = TestHelper.getFullPath("DarkMix3Workbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_3_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_3_STYLE, "A1"));
	}
	
	@Test
	public void darkMix4Workbook() {
		String excelFilename = TestHelper.getFullPath("DarkMix4Workbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_4_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_DARK_MIX_4_STYLE, "A1"));
	}
	
	
//	@Test
	public void darkWorkbook() {
		String excelFilename = TestHelper.getFullPath("DarkWorkbook.xlsx");
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
