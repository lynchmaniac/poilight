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
public class LightStyleTest {
	
	@Test
	public void lightGrayWorkbook() {
		String excelFilename = "d:\\tmp\\LightGrayWorkbook.xlsx";
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
		String excelFilename = "d:\\tmp\\LightBlueWorkbook.xlsx";
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
		String excelFilename = "d:\\tmp\\LightRedWorkbook.xlsx";
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
		String excelFilename = "d:\\tmp\\LightGreenWorkbook.xlsx";
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
		String excelFilename = "d:\\tmp\\LightPurpleWorkbook.xlsx";
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
		String excelFilename = "d:\\tmp\\LightTurquoiseWorkbook.xlsx";
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
		String excelFilename = "d:\\tmp\\LightOrangeWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "A1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "E1"));
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "I1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "I1"));
	}
	
	@Test
	public void lightWorkbook() {
		String excelFilename = "d:\\tmp\\LightWorkbook.xlsx";
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
