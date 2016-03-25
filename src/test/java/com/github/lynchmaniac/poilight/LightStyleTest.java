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
public class LightStyleTest {
	
	@BeforeClass
	public static void cleanRessources() {
		new File(TestHelper.getFullPath("LightGrayWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("LightBlueWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("LightRedWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("LightGreenWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("LightPurpleWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("LightTurquoiseWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("LightOrangeWorkbook.xlsx")).delete();
		new File(TestHelper.getFullPath("LightWorkbook.xlsx")).delete();
	}
	
	@Test
	public void lightGrayWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightGrayWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, "I1"));
	}
	
	@Test
	public void lightBlueWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightBlueWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, "I1"));
		
	}
	
	
	@Test
	public void lightRedWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightRedWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_3_STYLE, "I1"));
	}
	
	@Test
	public void lightGreenWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightGreenWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, "I1"));
	}
	
	@Test
	public void lightPurpleWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightPurpleWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, "I1"));
	}
	
	@Test
	public void lightTurquoiseWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightTurquoiseWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, "I1"));
	}
	
	@Test
	public void lightOrangeWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightOrangeWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "I1"));
	}
	
//	@Test
	public void lightWorkbook() {
		String excelFilename = TestHelper.getFullPath("LightWorkbook.xlsx");
		Workbook wb = new XSSFWorkbook();

		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, "B3"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, "B17"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, "B31"));
		
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, "F3"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, "F17"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, "F31"));

		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_1_STYLE, "J3"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_2_STYLE, "J17"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_3_STYLE, "J31"));

		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, "N3"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, "N17"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, "N31"));
		
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, "R3"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, "R17"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, "R31"));

		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, "V3"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, "V17"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, "V31"));

		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "Z3"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "Z17"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "Z31"));


		PoiLight.writeExcel(wb, excelFilename);
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, "B3"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, "B17"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, "B31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, "F3"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, "F17"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, "F31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_1_STYLE, "J3"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_2_STYLE, "J17"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_RED_3_STYLE, "J31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, "N3"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, "N17"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, "N31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, "R3"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, "R17"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, "R31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, "V3"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, "V17"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, "V31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "Z3"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "Z17"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "Z31"));
	}
}
