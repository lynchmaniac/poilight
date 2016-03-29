package com.github.lynchmaniac.poilight;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

import com.github.lynchmaniac.poilight.enumerations.BoardStyles;
import com.github.lynchmaniac.poilight.helpers.StyleHelper;
import com.github.lynchmaniac.poilight.models.ExcelCell;
import com.github.lynchmaniac.poilight.models.ExcelRow;
import com.github.lynchmaniac.poilight.models.Table;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;
import java.io.IOException;


/**
 * @author piard
 * @since 0.1
 *
 */
public class PoiLightTest {

  private static Table data = TestHelper.getTable();

  /**
   * Clean all the resources before the begining of the test.
   */
  @BeforeClass
  public static void cleanRessources() {
    new File(TestHelper.getFullPath("DefaultWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("DefaultStreamingWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("CustomStyleWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("AllStylesWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("DefaultStreamingWorkbook.xlsx")).delete();
  }

  @Test
  public void defaultWorkbook() {
    String excelFilename = TestHelper.getFullPath("DefaultWorkbook.xlsx");
    PoiLight.generateExcel(excelFilename, data);
    TestHelper.testTable(excelFilename, data);
  }

  @Test
  public void defaultStreamingWorkbook() throws IOException {
    String excelFilename = TestHelper.getFullPath("DefaultStreamingWorkbook.xlsx");
    PoiLight.generateStreamingExcel(excelFilename, data);
    TestHelper.testTable(excelFilename, data);
  }

  @Test
  public void customStyleWorkbook() {
    Workbook wb = new XSSFWorkbook();
    
    Table table = new Table();
    table.addHeader(new ExcelCell("ID"));
    table.addHeader(new ExcelCell("NOM"));
    table.addHeader(new ExcelCell("TITRE"));

    XSSFCellStyle cs = (XSSFCellStyle) wb.createCellStyle();
    cs.setFillForegroundColor(StyleHelper.getColor(128, 100, 162));
    cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
    table.setPosition("A1");
    table.addData(new ExcelRow(new ExcelCell("1"), new ExcelCell("Henri Loevenbruck"), new ExcelCell("L'apothicaire")));
    table.addData(new ExcelRow(new ExcelCell(2), new ExcelCell("Cyril Massarotto"), new ExcelCell("Dieu est un pote à moi")));
    table.addData(new ExcelRow(new ExcelCell(3), new ExcelCell("Bernard Werber"), new ExcelCell("Les fourmis")));
    table.addData(new ExcelRow(new ExcelCell(4, cs), new ExcelCell("Maxime Chattam"), new ExcelCell("In Tenebris")));
    table.addData(new ExcelRow(new ExcelCell(5), new ExcelCell("Franck Thilliez"), new ExcelCell("Pandemia")));

    PoiLight.createTable(wb, table);

    String excelFilename = TestHelper.getFullPath("CustomStyleWorkbook.xlsx");
    PoiLight.writeExcel(wb, excelFilename);

    try {
      wb = new XSSFWorkbook(excelFilename);
      Sheet sheet = wb.getSheet(PoiLightConst.DEFAULT_SHEET_NAME);
      Row row = sheet.getRow(4);
      Cell cell = row.getCell(0);
      XSSFCellStyle csValue = (XSSFCellStyle) cell.getCellStyle();
      assertEquals(CellStyle.SOLID_FOREGROUND, csValue.getFillPattern());
      assertEquals(StyleHelper.getColor(128, 100, 162), csValue.getFillForegroundXSSFColor()); 

      wb.close();
    } catch (IOException exception) {
      exception.printStackTrace();
      assertFalse(true);
    }

  }

  
  @Test
  public void tableNewStyleWorkbook() {
    Table table = new Table();
    table.addHeaders(new ExcelCell("ID"), new ExcelCell("NOM"), new ExcelCell("TITRE"), new ExcelCell("FORMULE"));
    table.setPosition("D4");
    table.setSheetName("test");
    table.setStyle(BoardStyles.BOARD_DARK_MIX_4_STYLE);
    table.addData(new ExcelRow(new ExcelCell(1), new ExcelCell(2), new ExcelCell(3), new ExcelCell("SUM(D5:F5)", true)));
    table.addData(new ExcelRow(new ExcelCell(2), new ExcelCell(10), new ExcelCell(5641), new ExcelCell("SUM(D6:F6)", true)));
    table.addData(new ExcelRow(new ExcelCell(3), new ExcelCell(20), new ExcelCell(654), new ExcelCell("SUM(D7:F7)", true)));
    table.addData(new ExcelRow(new ExcelCell(4), new ExcelCell(30), new ExcelCell(43), new ExcelCell("SUM(D8:F8)", true)));
    
    PoiLight.generateExcel(TestHelper.getFullPath("TableNewStyleWorkbook.xlsx"), table);
    
  }
  
  @Test
  public void tableNewStyle2Workbook() {
    Table table = new Table();
    table.addHeaders("ID", "NOM", "TITRE");
    table.setPosition("D4");
    table.setSheetName("test");
    table.setStyle(BoardStyles.BOARD_DARK_MIX_4_STYLE);
    table.addData(new ExcelRow(1, "Henri Loevenbruck", "L'apothicaire"));
    table.addData(new ExcelRow(2, "Cyril Massarotto", "Dieu est un pote à moi"));
    table.addData(new ExcelRow(3, "Bernard Werber", "Les fourmis"));
    table.addData(new ExcelRow(4, "Maxime Chattam", "In Tenebris"));
    table.addData(new ExcelRow(5, "Franck Thilliez", "Pandemia"));
    
    String excelFilename = TestHelper.getFullPath("TableNewStyle2Workbook.xlsx");
    PoiLight.generateExcel(excelFilename, table);
    TestHelper.testTable(excelFilename, table);
  }


  /**
   * Create all the predefined Style in the same Workbook with
   * 3 differents tabs, one light, one medium and one dark.
   */
  @Test
  public void allStylesWorkbook() {
    String excelFilename = TestHelper.getFullPath("AllStylesWorkbook.xlsx");
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
    try {
      wb.close();
    } catch (IOException exception) {
      System.out.println(exception.getMessage());
      assertFalse(true);
    }

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
