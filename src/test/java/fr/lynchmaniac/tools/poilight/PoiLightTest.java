/**
 * 
 */
package fr.lynchmaniac.tools.poilight;

import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import fr.lynchmaniac.tools.poilight.entite.CellContent;
import fr.lynchmaniac.tools.poilight.entite.RowContent;
import fr.lynchmaniac.tools.poilight.entite.Table;
import fr.lynchmaniac.tools.poilight.enumeration.BoardStyles;
import fr.lynchmaniac.tools.poilight.helpers.StyleHelper;

/**
 * @author piard
 * @since 0.1
 *
 */
public class PoiLightTest {


	private static Table data = TestHelper.getTable();
	
	@Test
	public void defaultWorkbook() {
		String excelFilename = "d:\\tmp\\DefaultWorkbook.xlsx";
		PoiLight.generateExcel(excelFilename, data);
		TestHelper.testTable(excelFilename, data);
	}
	
	@Test
	public void defaultStreamingWorkbook() throws IOException {
		String excelFilename = "d:\\tmp\\DefaultStreamingWorkbook.xlsx";
		PoiLight.generateStreamingExcel(excelFilename, data);
		TestHelper.testTable(excelFilename, data);
	}

	@Test
	public void customStyleWorkbook() {
		String excelFilename = "d:\\tmp\\CustomStyleWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		
		Table table = new Table();
		table.addHeader(new CellContent("ID"));
		table.addHeader(new CellContent("NOM"));
		table.addHeader(new CellContent("TITRE"));
		
		CellStyle cs = wb.createCellStyle();
		cs.setFillForegroundColor(StyleHelper.getColor(128, 100, 162).getIndex());
		cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		
		table.addData(new RowContent(new CellContent(4, cs), new CellContent("Maxime Chattam"), new CellContent("In Tenebris")));
		table.addData(new RowContent(new CellContent(5), new CellContent("Franck Thilliez"), new CellContent("Pandemia")));
		
		
		PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "A1"));
		PoiLight.writeExcel(wb, excelFilename);

		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "A1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "E1"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "I1"));
	}
	
	
	
	@Test
	public void AllStylesWorkbook() {
		String excelFilename = "d:\\tmp\\AllStylesWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();

		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, "B4"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, "B13"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, "B22"));
		
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, "F4"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, "F13"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, "F22"));

		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_RED_1_STYLE, "J4"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_RED_2_STYLE, "J13"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_RED_3_STYLE, "J22"));

		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, "N4"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, "N13"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, "N22"));
		
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, "R4"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, "R13"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, "R22"));

		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, "V4"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, "V13"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, "V22"));

		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "Z4"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "Z13"));
		PoiLight.createTable(wb, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "Z22"));
		
		
		
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, "B4"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, "B13"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, "B22"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, "B31"));
		
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, "F4"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, "F13"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, "F22"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, "F31"));

		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_1_STYLE, "J4"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_2_STYLE, "J13"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_3_STYLE, "J22"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_4_STYLE, "J31"));

		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, "N4"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, "N13"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, "N22"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, "N31"));
		
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, "R4"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, "R13"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, "R22"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, "R31"));

		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, "V4"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, "V13"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, "V22"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, "V31"));

		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, "Z4"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, "Z13"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, "Z22"));
		PoiLight.createTable(wb, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, "Z31"));
		
		
		
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "B4"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "F4"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_RED_1_STYLE, "J4"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "N4"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_PURPLE_1_STYLE, "R4"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, "V4"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_ORANGE_1_STYLE, "Z4"));
		
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_1_STYLE, "B13"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_2_STYLE, "F13"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_3_STYLE, "J13"));
		PoiLight.createTable(wb, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_4_STYLE, "N13"));


		PoiLight.writeExcel(wb, excelFilename);
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, "B4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, "B13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, "B22"));
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, "F4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, "F13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, "F22"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_RED_1_STYLE, "J4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_RED_2_STYLE, "J13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_RED_3_STYLE, "J22"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, "N4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, "N13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, "N22"));
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, "R4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, "R13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, "R22"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, "V4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, "V13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, "V22"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, "Z4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, "Z13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Light", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, "Z22"));
		
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, "B4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, "B13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, "B22"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, "B31"));
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, "F4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, "F13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, "F22"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, "F31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_1_STYLE, "J4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_2_STYLE, "J13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_3_STYLE, "J22"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_RED_4_STYLE, "J31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, "N4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, "N13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, "N22"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, "N31"));
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, "R4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, "R13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, "R22"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, "R31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, "V4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, "V13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, "V22"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, "V31"));

		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, "Z4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, "Z13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, "Z22"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Medium", BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, "Z31"));
		
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_GRAY_1_STYLE, "B4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_BLUE_1_STYLE, "F4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_RED_1_STYLE, "J4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_GREEN_1_STYLE, "N4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_PURPLE_1_STYLE, "R4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, "V4"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_ORANGE_1_STYLE, "Z4"));
		
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_1_STYLE, "B13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_2_STYLE, "F13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_3_STYLE, "J13"));
		TestHelper.testTable(excelFilename, TestHelper.getTable("Dark", BoardStyles.BOARD_DARK_MIX_4_STYLE, "N13"));
	}
}
