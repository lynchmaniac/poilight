package com.github.lynchmaniac.poilight;

import com.github.lynchmaniac.poilight.PoiLight;
import com.github.lynchmaniac.poilight.enumerations.BoardStyles;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.File;

/**
 * @author piard
 * @since 0.1
 *
 */
public class MediumStyleTest {

  /**
   * Clean all the resources before the begining of the test.
   */
  @BeforeClass
  public static void cleanRessources() {
    new File(TestHelper.getFullPath("MediumGrayWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("MediumBlueWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("MediumRedWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("MediumGreenWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("MediumPurpleWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("MediumTurquoiseWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("MediumOrangeWorkbook.xlsx")).delete();
    new File(TestHelper.getFullPath("MediumWorkbook.xlsx")).delete();
  }

  @Test
  public void mediumGrayWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumGrayWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, "A1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, "E1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, "I1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, "M1"));
    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, "A1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, "E1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, "I1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, "M1"));
  }

  @Test
  public void mediumBlueWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumBlueWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, "A1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, "E1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, "I1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, "M1"));
    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, "A1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, "E1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, "I1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, "M1"));

  }


  @Test
  public void mediumRedWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumRedWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_1_STYLE, "A1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_2_STYLE, "E1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_3_STYLE, "I1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_4_STYLE, "M1"));
    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_1_STYLE, "A1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_2_STYLE, "E1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_3_STYLE, "I1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_4_STYLE, "M1"));
  }

  @Test
  public void mediumGreenWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumGreenWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, "A1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, "E1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, "I1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, "M1"));
    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, "A1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, "E1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, "I1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, "M1"));
  }

  @Test
  public void mediumPurpleWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumPurpleWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, "A1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, "E1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, "I1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, "M1"));
    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, "A1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, "E1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, "I1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, "M1"));
  }

  @Test
  public void mediumTurquoiseWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumTurquoiseWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, "A1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, "E1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, "I1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, "M1"));
    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, "A1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, "E1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, "I1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, "M1"));
  }

  @Test
  public void mediumOrangeWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumOrangeWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, "A1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, "E1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, "I1"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, "M1"));
    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, "A1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, "E1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, "I1"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, "M1"));
  }


  @Test
  public void mediumWorkbook() {
    String excelFilename = TestHelper.getFullPath("MediumWorkbook.xlsx");
    Workbook wb = new XSSFWorkbook();


    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, "B2"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, "B14"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, "B26"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, "B38"));

    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, "F2"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, "F14"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, "F26"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, "F38"));

    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_1_STYLE, "J2"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_2_STYLE, "J14"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_3_STYLE, "J26"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_4_STYLE, "J38"));

    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, "N2"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, "N14"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, "N26"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, "N38"));

    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, "R2"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, "R14"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, "R26"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, "R38"));

    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, "V2"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, "V14"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, "V26"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, "V38"));

    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, "Z2"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, "Z14"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, "Z26"));
    PoiLight.createTable(wb, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, "Z38"));

    PoiLight.writeExcel(wb, excelFilename);

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, "B2"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, "B14"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, "B26"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, "B38"));

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, "F2"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, "F14"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, "F26"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, "F38"));

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_1_STYLE, "J2"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_2_STYLE, "J14"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_3_STYLE, "J26"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_RED_4_STYLE, "J38"));

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, "N2"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, "N14"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, "N26"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, "N38"));

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, "R2"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, "R14"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, "R26"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, "R38"));

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, "V2"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, "V14"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, "V26"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, "V38"));

    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, "Z2"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, "Z14"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, "Z26"));
    TestHelper.testTable(excelFilename, TestHelper.getTable("custom", BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, "Z38"));
  }
}
