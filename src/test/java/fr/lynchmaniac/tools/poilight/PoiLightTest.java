/**
 * 
 */
package fr.lynchmaniac.tools.poilight;

import java.util.LinkedHashMap;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import fr.lynchmaniac.tools.poilight.entite.CellContent;
import fr.lynchmaniac.tools.poilight.entite.RowContent;
import fr.lynchmaniac.tools.poilight.enumeration.BoardStyles;

/**
 * @author piard
 * @since 0.1
 *
 */
public class PoiLightTest {
	
	private static LinkedHashMap<Integer, RowContent> data = getData();

	
	@Test
	public void defaultWorkbook() {
		String excelFilename = "d:\\tmp\\DefaultWorkbook.xlsx";
		PoiLight.generateExcel(excelFilename, data);
	}
	
	@Test
	public void defaultStreamingWorkbook() {
		String excelFilename = "d:\\tmp\\DefaultStreamingWorkbook.xlsx";
		PoiLight.generateStreamingExcel(excelFilename, data);
	}
		
	@Test
	public void customBlueWorkbook() {
		String excelFilename = "d:\\tmp\\CustomBlueWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, 1, 1);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, 1, 5);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, 1, 9);
		PoiLight.writeExcel(wb, excelFilename);
	}
	@Test
	public void customGrayWorkbook() {
		String excelFilename = "d:\\tmp\\CustomGrayWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, 1, 1);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, 1, 5);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, 1, 9);
		PoiLight.writeExcel(wb, excelFilename);
	}
	@Test
	public void customRedWorkbook() {
		String excelFilename = "d:\\tmp\\CustomRedWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_RED_1_STYLE, 1, 1);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_RED_2_STYLE, 1, 5);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_RED_3_STYLE, 1, 9);
		PoiLight.writeExcel(wb, excelFilename);
	}
	@Test
	public void customOrangeWorkbook() {
		String excelFilename = "d:\\tmp\\CustomOrangeWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, 1, 1);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, 17, 5);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, 31, 9);
		PoiLight.writeExcel(wb, excelFilename);
	}
	@Test
	public void lightWorkbook() {
		String excelFilename = "d:\\tmp\\LightWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, 3, 2);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, 17, 2);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, 31, 2);
		
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, 3, 6);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, 17, 6);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, 31, 6);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_RED_1_STYLE, 3, 10);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_RED_2_STYLE, 17, 10);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_RED_3_STYLE, 31, 10);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, 3, 14);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, 17, 14);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, 31, 14);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, 3, 18);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, 17, 18);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, 31, 18);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, 3, 22);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, 17, 22);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, 31, 22);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, 3, 26);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, 17, 26);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, 31, 26);


		PoiLight.writeExcel(wb, excelFilename);
	}
	@Test
	public void mediumWorkbook() {
		String excelFilename = "d:\\tmp\\MediumWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, 2, 2);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, 14, 2);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, 26, 2);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, 38, 2);
		
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, 2, 6);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, 14, 6);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, 26, 6);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, 38, 6);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_RED_1_STYLE, 2, 10);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_RED_2_STYLE, 14, 10);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_RED_3_STYLE, 26, 10);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_RED_4_STYLE, 38, 10);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, 2, 14);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, 14, 14);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, 26, 14);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, 38, 14);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, 2, 18);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, 14, 18);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, 26, 18);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, 38, 18);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, 2, 22);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, 14, 22);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, 26, 22);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, 38, 22);

		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, 2, 26);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, 14, 26);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, 26, 26);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, 38, 26);


		PoiLight.writeExcel(wb, excelFilename);
	}
	
	@Test
	public void darkWorkbook() {
		String excelFilename = "d:\\tmp\\DarkWorkbook.xlsx";
		Workbook wb = new XSSFWorkbook();
		
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_GRAY_1_STYLE, 2, 2);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_BLUE_1_STYLE, 2, 6);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_RED_1_STYLE, 2, 10);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_GREEN_1_STYLE, 2, 14);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_PURPLE_1_STYLE, 2, 18);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, 2, 22);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_ORANGE_1_STYLE, 2, 26);
		
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_MIX_1_STYLE, 15, 2);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_MIX_2_STYLE, 15, 6);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_MIX_3_STYLE, 15, 10);
		PoiLight.createSheet(wb, data, false, "custom", BoardStyles.BOARD_DARK_MIX_4_STYLE, 15, 14);


		PoiLight.writeExcel(wb, excelFilename);
	}
	
	
	private static LinkedHashMap<Integer, RowContent> getData() {
		LinkedHashMap<Integer, RowContent> hashMap = new LinkedHashMap<Integer, RowContent>();
		
		RowContent rowContent = new RowContent();
		rowContent.addValue(new CellContent("ID"));
		rowContent.addValue(new CellContent("NOM"));
		rowContent.addValue(new CellContent("TITRE"));
		hashMap.put(0, rowContent);

		rowContent = new RowContent(); 
		rowContent.addValue(new CellContent(1));
		rowContent.addValue(new CellContent("Henri Loevenbruck"));
		rowContent.addValue(new CellContent("L'apothicaire"));
		hashMap.put(1, rowContent);
		
		rowContent = new RowContent(); 
		rowContent.addValue(new CellContent(2));
		rowContent.addValue(new CellContent("Cyril Massarotto"));
		rowContent.addValue(new CellContent("Dieu est un pote à moi"));
		hashMap.put(2, rowContent);

		rowContent = new RowContent(); 
		rowContent.addValue(new CellContent(3));
		rowContent.addValue(new CellContent("Bernard Werber"));
		rowContent.addValue(new CellContent("Les fourmis"));
		hashMap.put(3, rowContent);

		rowContent = new RowContent(); 
		rowContent.addValue(new CellContent(4));
		rowContent.addValue(new CellContent("Maxime Chattam"));
		rowContent.addValue(new CellContent("In Tenebris"));
		hashMap.put(4, rowContent);

		rowContent = new RowContent(); 
		rowContent.addValue(new CellContent(5));
		rowContent.addValue(new CellContent("Franck Thilliez"));
		rowContent.addValue(new CellContent("Pandemia"));
		hashMap.put(5, rowContent);

		return hashMap;
	}
}
