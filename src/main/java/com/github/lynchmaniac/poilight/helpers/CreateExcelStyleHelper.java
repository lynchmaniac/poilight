package com.github.lynchmaniac.poilight.helpers;


import com.github.lynchmaniac.poilight.enumerations.BoardStyles;
import com.github.lynchmaniac.poilight.models.BorderStyle;
import com.github.lynchmaniac.poilight.models.TableStyle;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.HashMap;



/**
 * Create and store all the Excel's style. It's based on the Execl 2016 version.
 * You can use this style later with the hashmap styles.
 * 
 * @author vpiard
 * @since 0.1
 */
public class CreateExcelStyleHelper {

  /**
   * Stores all Excel's styles.
   */
  private static HashMap<BoardStyles, HashMap<String, TableStyle>> styles = null;

  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_MEDIUM<br>
   * Top : BORDER_THIN<br>
   * Left : BORDER_THIN<br>
   * Right : BORDER_THIN<br>
   */
  private static BorderStyle mediumThin = new BorderStyle(CellStyle.BORDER_MEDIUM, 
                                                          CellStyle.BORDER_THIN, 
                                                          CellStyle.BORDER_THIN,
                                                          CellStyle.BORDER_THIN);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_MEDIUM<br>
   * Top : BORDER_MEDIUM<br>
   * Left : BORDER_NONE<br>
   * Right : BORDER_NONE<br>
   */
  private static BorderStyle mediumTopBottom = new BorderStyle(CellStyle.BORDER_MEDIUM, 
                                                                CellStyle.BORDER_MEDIUM, 
                                                                CellStyle.BORDER_NONE, 
                                                                CellStyle.BORDER_NONE);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_MEDIUM<br>
   * Top : BORDER_NONE<br>
   * Left : BORDER_NONE<br>
   * Right : BORDER_NONE<br>
   */
  private static BorderStyle mediumBottom = new BorderStyle(CellStyle.BORDER_MEDIUM, 
                                                            CellStyle.BORDER_NONE, 
                                                            CellStyle.BORDER_NONE, 
                                                            CellStyle.BORDER_NONE);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_NONE<br>
   * Top : BORDER_NONE<br>
   * Left : BORDER_NONE<br>
   * Right : BORDER_NONE<br>
   */
  private static BorderStyle none = new BorderStyle(CellStyle.BORDER_NONE, 
                                                    CellStyle.BORDER_NONE, 
                                                    CellStyle.BORDER_NONE, 
                                                    CellStyle.BORDER_NONE);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_THIN<br>
   * Top : BORDER_NONE<br>
   * Left : BORDER_NONE<br>
   * Right : BORDER_NONE<br>
   */
  private static BorderStyle oneThin = new BorderStyle(CellStyle.BORDER_THIN, 
                                                        CellStyle.BORDER_NONE, 
                                                        CellStyle.BORDER_NONE, 
                                                        CellStyle.BORDER_NONE);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_THIN<br>
   * Top : BORDER_THIN<br>
   * Left : BORDER_NONE<br>
   * Right : BORDER_NONE<br>
   */
  private static BorderStyle noneThin = new BorderStyle(CellStyle.BORDER_THIN, 
                                                        CellStyle.BORDER_THIN, 
                                                        CellStyle.BORDER_NONE, 
                                                        CellStyle.BORDER_NONE);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_THIN<br>
   * Top : BORDER_THIN<br>
   * Left : BORDER_THIN<br>
   * Right : BORDER_THIN<br>
   */
  private static BorderStyle allThin = new BorderStyle(CellStyle.BORDER_THIN, 
                                                        CellStyle.BORDER_THIN, 
                                                        CellStyle.BORDER_THIN, 
                                                        CellStyle.BORDER_THIN);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_MEDIUM<br>
   * Top : BORDER_MEDIUM<br>
   * Left : BORDER_MEDIUM<br>
   * Right : BORDER_MEDIUM<br>
   */
  private static BorderStyle allMedium = new BorderStyle(CellStyle.BORDER_MEDIUM, 
                                                          CellStyle.BORDER_MEDIUM, 
                                                          CellStyle.BORDER_MEDIUM, 
                                                          CellStyle.BORDER_MEDIUM);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_THIN<br>
   * Top : BORDER_THIN<br>
   * Left : BORDER_MEDIUM<br>
   * Right : BORDER_MEDIUM<br>
   */
  private static BorderStyle mediumLeftRight = new BorderStyle(CellStyle.BORDER_THIN, 
                                                                CellStyle.BORDER_THIN, 
                                                                CellStyle.BORDER_MEDIUM, 
                                                                CellStyle.BORDER_MEDIUM);
  /**
   * The style of the border of the cell.
   * <br>
   * Bottom : BORDER_MEDIUM<br>
   * Top : BORDER_THIN<br>
   * Left : BORDER_MEDIUM<br>
   * Right : BORDER_MEDIUM<br>
   */
  private static BorderStyle mediumBottomLeftRight = new BorderStyle(CellStyle.BORDER_MEDIUM, 
                                                                      CellStyle.BORDER_THIN, 
                                                                      CellStyle.BORDER_MEDIUM, 
                                                                      CellStyle.BORDER_MEDIUM);


  /**
   * The font color WHITE : 255, 255, 255.
   */
  private static XSSFColor FONT_WHITE = StyleHelper.getColor(255, 255, 255);
  /**
   * The font color BLACK : 0, 0, 0.
   */
  private static XSSFColor FONT_BLACK = StyleHelper.getColor(0, 0, 0);
  /**
   * The font color BLUE : 54, 96, 146.
   */
  private static XSSFColor FONT_BLUE = StyleHelper.getColor(54, 96, 146);
  /**
   * The font color RED : 150, 54, 52.
   */
  private static XSSFColor FONT_RED = StyleHelper.getColor(150, 54, 52);
  /**
   * The font color GREEN : 118, 147, 60.
   */
  private static XSSFColor FONT_GREEN = StyleHelper.getColor(118, 147, 60);
  /**
   * The font color PURPLE : 96, 73, 122.
   */
  private static XSSFColor FONT_PURPLE = StyleHelper.getColor(96, 73, 122);
  /**
   * The font color TURQUOISE : 49, 134, 155.
   */
  private static XSSFColor FONT_TURQUOISE = StyleHelper.getColor(49, 134, 155);
  /**
   * The font color ORANGE : 226, 107,10.
   */
  private static XSSFColor FONT_ORANGE = StyleHelper.getColor(226, 107,10);


  /**
   * Color WHITE : 255, 255, 255.
   */
  private static XSSFColor BACKGROUND_WHITE = StyleHelper.getColor(255, 255, 255);
  /**
   * Color BLACK : 0, 0, 0.
   */
  private static XSSFColor BACKGROUND_BLACK = StyleHelper.getColor(0, 0, 0);
  /**
   * Color GRAY : 217, 217, 217.
   */
  private static XSSFColor BACKGROUND_GRAY = StyleHelper.getColor(217, 217, 217);
  /**
   * Color GRAY Medium : 166, 166, 166.
   */
  private static XSSFColor BACKGROUND_GRAY_MEDIUM = StyleHelper.getColor(166, 166, 166);
  /**
   * Color GRAY DARK : 115, 115, 115.
   */
  private static XSSFColor BACKGROUND_GRAY_DARK = StyleHelper.getColor(115, 115, 115);
  /**
   * Color GRAY DARKER : 64, 64, 64.
   */
  private static XSSFColor BACKGROUND_GRAY_DARKER = StyleHelper.getColor(64, 64, 64);
  /**
   * Color BLUE : 220, 230,241.
   */
  private static XSSFColor BACKGROUND_BLUE = StyleHelper.getColor(220, 230,241);
  /**
   * Color BLUE : 79, 129, 189.
   */
  private static XSSFColor BACKGROUND_BLUE_HEAD = StyleHelper.getColor(79, 129, 189);
  /**
   * Color BLUE MEDIUM: 184, 204, 228.
   */
  private static XSSFColor BACKGROUND_BLUE_MEDIUM = StyleHelper.getColor(184, 204, 228);
  /**
   * Color BLUE DARK : 79, 129, 189.
   */
  private static XSSFColor BACKGROUND_BLUE_DARK = StyleHelper.getColor(79, 129, 189);
  /**
   * Color BLUE DARKER : 54, 96, 146.
   */
  private static XSSFColor BACKGROUND_BLUE_DARKER = StyleHelper.getColor(54, 96, 146);
  /**
   * Color RED : 242, 220, 219.
   */
  private static XSSFColor BACKGROUND_RED = StyleHelper.getColor(242, 220, 219);
  /**
   * Color RED HEAD : 192, 80, 77.
   */
  private static XSSFColor BACKGROUND_RED_HEAD = StyleHelper.getColor(192, 80, 77);
  /**
   * Color RED MEDIUM : 230, 184, 183.
   */
  private static XSSFColor BACKGROUND_RED_MEDIUM = StyleHelper.getColor(230, 184, 183);
  /**
   * Color RED DARK : 192, 80, 77.
   */
  private static XSSFColor BACKGROUND_RED_DARK = StyleHelper.getColor(192, 80, 77);
  /**
   * Color RED DARKER : 150, 54, 52.
   */
  private static XSSFColor BACKGROUND_RED_DARKER = StyleHelper.getColor(150, 54, 52);
  /**
   * Color GREEN : 235, 241, 222.
   */
  private static XSSFColor BACKGROUND_GREEN = StyleHelper.getColor(235, 241, 222);
  /**
   * Color GREEN HEAD : 155, 187, 89.
   */
  private static XSSFColor BACKGROUND_GREEN_HEAD = StyleHelper.getColor(155, 187, 89);
  /**
   * Color GREEN MEDIUM : 216, 228, 188.
   */
  private static XSSFColor BACKGROUND_GREEN_MEDIUM = StyleHelper.getColor(216, 228, 188);
  /**
   * Color GREEN DARK : 155, 187, 89.
   */
  private static XSSFColor BACKGROUND_GREEN_DARK = StyleHelper.getColor(155, 187, 89);
  /**
   * Color GREEN DARKER : 118, 147, 60.
   */
  private static XSSFColor BACKGROUND_GREEN_DARKER = StyleHelper.getColor(118, 147, 60);
  /**
   * Color PURPLE : 228, 223, 236.
   */
  private static XSSFColor BACKGROUND_PURPLE = StyleHelper.getColor(228, 223, 236);
  /**
   * Color PURPLE HEAD : 128, 100, 162.
   */
  private static XSSFColor BACKGROUND_PURPLE_HEAD = StyleHelper.getColor(128, 100, 162);
  /**
   * Color PURPLE MEDIUM : 204, 192, 218.
   */
  private static XSSFColor BACKGROUND_PURPLE_MEDIUM = StyleHelper.getColor(204, 192, 218);
  /**
   * Color PURPLE DARK : 128, 100, 162.
   */
  private static XSSFColor BACKGROUND_PURPLE_DARK = StyleHelper.getColor(128, 100, 162);
  /**
   * Color PURPLE DARKER : 96, 73, 122.
   */
  private static XSSFColor BACKGROUND_PURPLE_DARKER = StyleHelper.getColor(96, 73, 122);
  /**
   * Color TURQUOISE  : 218, 238, 243.
   */
  private static XSSFColor BACKGROUND_TURQUOISE = StyleHelper.getColor(218, 238, 243);
  /**
   * Color TURQUOISE HEAD : 75, 172, 198.
   */
  private static XSSFColor BACKGROUND_TURQUOISE_HEAD = StyleHelper.getColor(75, 172, 198);
  /**
   * Color TURQUOISE MEDIUM : 183, 222, 232.
   */
  private static XSSFColor BACKGROUND_TURQUOISE_MEDIUM = StyleHelper.getColor(183, 222, 232);
  /**
   * Color TURQUOISE DARK : 75, 172, 198.
   */
  private static XSSFColor BACKGROUND_TURQUOISE_DARK = StyleHelper.getColor(75, 172, 198);
  /**
   * Color TURQUOISE DARKER : 49, 134, 155.
   */
  private static XSSFColor BACKGROUND_TURQUOISE_DARKER = StyleHelper.getColor(49, 134, 155);
  /**
   * Color ORANGE : 253, 233, 217.
   */
  private static XSSFColor BACKGROUND_ORANGE = StyleHelper.getColor(253, 233, 217);
  /**
   * Color ORANGE HEAD : 247, 150, 70.
   */
  private static XSSFColor BACKGROUND_ORANGE_HEAD = StyleHelper.getColor(247, 150, 70);
  /**
   * Color ORANGE MEDIUM : 252, 213, 180.
   */
  private static XSSFColor BACKGROUND_ORANGE_MEDIUM = StyleHelper.getColor(252, 213, 180);
  /**
   * Color ORANGE DARK : 247, 150, 70.
   */
  private static XSSFColor BACKGROUND_ORANGE_DARK = StyleHelper.getColor(247, 150, 70);
  /**
   * Color ORANGE DARKER : 226, 107, 10.
   */
  private static XSSFColor BACKGROUND_ORANGE_DARKER = StyleHelper.getColor(226, 107, 10);



  /**
   * Border WHITE : 255, 255, 255.
   */
  private static XSSFColor BORDER_WHITE = StyleHelper.getColor(255, 255, 255);
  /**
   * Border BLACK : 0, 0, 0.
   */
  private static XSSFColor BORDER_BLACK = StyleHelper.getColor(0, 0, 0);
  /**
   * Border BLUE : 79, 129, 189.
   */
  private static XSSFColor BORDER_BLUE = StyleHelper.getColor(79, 129, 189);
  /**
   * Border RED : 192, 80, 77.
   */
  private static XSSFColor BORDER_RED = StyleHelper.getColor(192, 80, 77);
  /**
   * Border GREEN : 155, 187, 89.
   */
  private static XSSFColor BORDER_GREEN = StyleHelper.getColor(155, 187, 89);
  /**
   * Border PURPLE : 128, 100, 162.
   */
  private static XSSFColor BORDER_PURPLE = StyleHelper.getColor(128, 100, 162);
  /**
   * Border TURQUOISE : 75, 172, 198.
   */
  private static XSSFColor BORDER_TURQUOISE = StyleHelper.getColor(75, 172, 198);
  /**
   * Border ORANGE : 247, 150, 70.
   */
  private static XSSFColor BORDER_ORANGE = StyleHelper.getColor(247, 150, 70);



  /**
   * Generate the referential of the Excel Styles.
   * 
   * @return the excel styles
   */
  public static HashMap<BoardStyles, HashMap<String, TableStyle>> getExcelStyle() {
    if (styles == null) {
      initializeStyles();
    }
    return styles;
  }


  /**
   * Initialize all the Excel Style from Excel 2016.
   */
  private static void initializeStyles() {
    styles = new  HashMap<BoardStyles, HashMap<String, TableStyle>>();
    createLightStyles();
    createMediumStyles();
    createDarkStyles();
  }

  /**
   * Configure the light style.
   */
  private static void createLightStyles() {

    HashMap<String, TableStyle> currentStyle = new HashMap<String, TableStyle>();
    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_BLACK, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_GRAY_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_BLACK, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_BLACK, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_GRAY_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_BLACK, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, allThin));
    styles.put(BoardStyles.BOARD_LIGHT_GRAY_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLUE, BACKGROUND_WHITE, BORDER_BLUE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLUE, BORDER_BLUE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLUE, BORDER_BLUE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLUE, BORDER_BLUE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLUE, BORDER_BLUE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_BLUE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLUE_HEAD, BORDER_BLUE, none));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_BLUE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_BLUE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_BLUE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_BLUE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, allThin));
    styles.put(BoardStyles.BOARD_LIGHT_BLUE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_RED, BACKGROUND_WHITE, BORDER_RED, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_RED, FONT_RED, BORDER_RED, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_RED, FONT_RED, BORDER_RED, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_RED, FONT_RED, BORDER_RED, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_RED, FONT_RED, BORDER_RED, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_RED_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_RED_HEAD, BORDER_RED, none));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_RED, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_RED, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_RED_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_RED, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, allThin));
    styles.put(BoardStyles.BOARD_LIGHT_RED_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_GREEN, BACKGROUND_WHITE, BORDER_GREEN, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_GREEN, BORDER_GREEN, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_GREEN, BORDER_GREEN, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_GREEN, BORDER_GREEN, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_GREEN, BORDER_GREEN, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_GREEN_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_GREEN_HEAD, BORDER_GREEN, none));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_GREEN, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_GREEN, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_GREEN_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_GREEN, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, allThin));
    styles.put(BoardStyles.BOARD_LIGHT_GREEN_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_PURPLE, BACKGROUND_WHITE, BORDER_PURPLE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_PURPLE, BORDER_PURPLE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_PURPLE, BORDER_PURPLE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_PURPLE, BORDER_PURPLE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_PURPLE, BORDER_PURPLE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_PURPLE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_PURPLE_HEAD, BORDER_PURPLE, none));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_PURPLE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_PURPLE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_PURPLE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_PURPLE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, allThin));
    styles.put(BoardStyles.BOARD_LIGHT_PURPLE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_TURQUOISE, BACKGROUND_WHITE, BORDER_TURQUOISE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_TURQUOISE, BORDER_TURQUOISE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_TURQUOISE, BORDER_TURQUOISE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_TURQUOISE, BORDER_TURQUOISE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_TURQUOISE, BORDER_TURQUOISE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_TURQUOISE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_TURQUOISE_HEAD, BORDER_TURQUOISE, none));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_TURQUOISE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_TURQUOISE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_TURQUOISE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_TURQUOISE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, allThin));
    styles.put(BoardStyles.BOARD_LIGHT_TURQUOISE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_ORANGE, BACKGROUND_WHITE, BORDER_ORANGE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_ORANGE, BORDER_ORANGE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_ORANGE, BORDER_ORANGE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_ORANGE, BORDER_ORANGE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_ORANGE, BORDER_ORANGE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_ORANGE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_ORANGE_HEAD, BORDER_ORANGE, none));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_ORANGE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, null, FONT_BLACK, BORDER_ORANGE, oneThin));
    styles.put(BoardStyles.BOARD_LIGHT_ORANGE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_WHITE, BORDER_ORANGE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, allThin));
    styles.put(BoardStyles.BOARD_LIGHT_ORANGE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_GRAY_MEDIUM, BORDER_BLACK, allMedium));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_WHITE, FONT_BLACK, BORDER_BLACK, mediumLeftRight));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_WHITE, FONT_BLACK, BORDER_BLACK, mediumBottomLeftRight));
    styles.put(BoardStyles.BOARD_DEFAULT_STYLE, currentStyle);
  }

  /**
   * Configure the medium style.
   */
  private static void createMediumStyles() {

    HashMap<String, TableStyle> currentStyle = new HashMap<String, TableStyle>();
    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_BLACK, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, oneThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, oneThin));
    styles.put(BoardStyles.BOARD_MEDIUM_GRAY_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_GRAY_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_BLACK, mediumTopBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumThin));
    styles.put(BoardStyles.BOARD_MEDIUM_GRAY_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_GRAY, BORDER_BLACK, allThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_GRAY_4_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLUE_HEAD, BORDER_BLUE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, oneThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_BLUE, FONT_BLACK, BORDER_BLUE, oneThin));
    styles.put(BoardStyles.BOARD_MEDIUM_BLUE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLUE_HEAD, BORDER_WHITE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_BLUE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLUE_HEAD, BORDER_BLACK, mediumTopBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    styles.put(BoardStyles.BOARD_MEDIUM_BLUE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_BLUE, BORDER_BLUE, allThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLUE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLUE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLUE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLUE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_BLUE_4_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_RED_HEAD, BORDER_RED, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, oneThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_RED, FONT_BLACK, BORDER_RED, oneThin));
    styles.put(BoardStyles.BOARD_MEDIUM_RED_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_RED_HEAD, BORDER_WHITE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_RED_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_RED_HEAD, BORDER_BLACK, mediumTopBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    styles.put(BoardStyles.BOARD_MEDIUM_RED_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_RED, BORDER_RED, allThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_RED, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_RED, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_RED, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_RED, BACKGROUND_RED_MEDIUM, FONT_BLACK, BORDER_RED, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_RED_4_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_GREEN_HEAD, BORDER_GREEN, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, oneThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GREEN, FONT_BLACK, BORDER_GREEN, oneThin));
    styles.put(BoardStyles.BOARD_MEDIUM_GREEN_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_GREEN_HEAD, BORDER_WHITE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_GREEN_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_GREEN_HEAD, BORDER_BLACK, mediumTopBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    styles.put(BoardStyles.BOARD_MEDIUM_GREEN_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_GREEN, BORDER_GREEN, allThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_GREEN, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_GREEN, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_GREEN, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_GREEN, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_GREEN_4_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_PURPLE_HEAD, BORDER_PURPLE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, oneThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_PURPLE, FONT_BLACK, BORDER_PURPLE, oneThin));
    styles.put(BoardStyles.BOARD_MEDIUM_PURPLE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_PURPLE_HEAD, BORDER_WHITE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_PURPLE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_PURPLE_HEAD, BORDER_BLACK, mediumTopBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    styles.put(BoardStyles.BOARD_MEDIUM_PURPLE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_PURPLE, BORDER_PURPLE, allThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_PURPLE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_PURPLE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_PURPLE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_PURPLE, BACKGROUND_PURPLE_MEDIUM, FONT_BLACK, BORDER_PURPLE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_PURPLE_4_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_TURQUOISE_HEAD, BORDER_TURQUOISE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, oneThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_TURQUOISE, FONT_BLACK, BORDER_TURQUOISE, oneThin));
    styles.put(BoardStyles.BOARD_MEDIUM_TURQUOISE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_TURQUOISE_HEAD, BORDER_WHITE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_TURQUOISE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_TURQUOISE_HEAD, BORDER_BLACK, mediumTopBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    styles.put(BoardStyles.BOARD_MEDIUM_TURQUOISE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_TURQUOISE, BORDER_TURQUOISE, allThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_TURQUOISE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_TURQUOISE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_TURQUOISE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_TURQUOISE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_TURQUOISE_4_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_ORANGE_HEAD, BORDER_ORANGE, noneThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, oneThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, oneThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, oneThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_ORANGE, FONT_BLACK, BORDER_ORANGE, oneThin));
    styles.put(BoardStyles.BOARD_MEDIUM_ORANGE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_ORANGE_HEAD, BORDER_WHITE, mediumThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_WHITE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_ORANGE_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_ORANGE_HEAD, BORDER_BLACK, mediumTopBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_WHITE, BACKGROUND_GRAY, FONT_BLACK, BORDER_BLACK, mediumBottom));
    styles.put(BoardStyles.BOARD_MEDIUM_ORANGE_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_BLACK, BACKGROUND_ORANGE, BORDER_ORANGE, allThin));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_ORANGE, allThin));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_ORANGE, allThin));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_ORANGE, allThin));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_ORANGE, BACKGROUND_ORANGE_MEDIUM, FONT_BLACK, BORDER_ORANGE, allThin));
    styles.put(BoardStyles.BOARD_MEDIUM_ORANGE_4_STYLE, currentStyle);
  }

  /**
   * Configure the dark style.
   */
  private static void createDarkStyles() {
    HashMap<String, TableStyle> currentStyle = new HashMap<String, TableStyle>();
    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GRAY_DARK, BACKGROUND_GRAY_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GRAY_DARK, BACKGROUND_GRAY_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GRAY_DARK, BACKGROUND_GRAY_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GRAY_DARK, BACKGROUND_GRAY_DARKER, FONT_WHITE, BORDER_WHITE, none));
    styles.put(BoardStyles.BOARD_DARK_GRAY_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_BLUE_DARK, BACKGROUND_BLUE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_BLUE_DARK, BACKGROUND_BLUE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_BLUE_DARK, BACKGROUND_BLUE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_BLUE_DARK, BACKGROUND_BLUE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    styles.put(BoardStyles.BOARD_DARK_BLUE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_RED_DARK, BACKGROUND_RED_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_RED_DARK, BACKGROUND_RED_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_RED_DARK, BACKGROUND_RED_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_RED_DARK, BACKGROUND_RED_DARKER, FONT_WHITE, BORDER_WHITE, none));
    styles.put(BoardStyles.BOARD_DARK_RED_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GREEN_DARK, BACKGROUND_GREEN_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GREEN_DARK, BACKGROUND_GREEN_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GREEN_DARK, BACKGROUND_GREEN_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GREEN_DARK, BACKGROUND_GREEN_DARKER, FONT_WHITE, BORDER_WHITE, none));
    styles.put(BoardStyles.BOARD_DARK_GREEN_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_PURPLE_DARK, BACKGROUND_PURPLE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_PURPLE_DARK, BACKGROUND_PURPLE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_PURPLE_DARK, BACKGROUND_PURPLE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_PURPLE_DARK, BACKGROUND_PURPLE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    styles.put(BoardStyles.BOARD_DARK_PURPLE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE_DARK, BACKGROUND_TURQUOISE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE_DARK, BACKGROUND_TURQUOISE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE_DARK, BACKGROUND_TURQUOISE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE_DARK, BACKGROUND_TURQUOISE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    styles.put(BoardStyles.BOARD_DARK_TURQUOISE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentStyle.put("HEAD", getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_WHITE, mediumBottom));
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_ORANGE_DARK, BACKGROUND_ORANGE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_ORANGE_DARK, BACKGROUND_ORANGE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_ORANGE_DARK, BACKGROUND_ORANGE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_ORANGE_DARK, BACKGROUND_ORANGE_DARKER, FONT_WHITE, BORDER_WHITE, none));
    styles.put(BoardStyles.BOARD_DARK_ORANGE_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    TableStyle currentBs = getHeaderStyle(FONT_WHITE, BACKGROUND_BLACK, BORDER_BLACK, none);
    currentBs.setBold(false);
    currentStyle.put("HEAD", currentBs);
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GRAY, BACKGROUND_GRAY_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    styles.put(BoardStyles.BOARD_DARK_MIX_1_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentBs = getHeaderStyle(FONT_WHITE, BACKGROUND_RED_HEAD, BORDER_BLACK, none);
    currentBs.setBold(false);
    currentStyle.put("HEAD", currentBs);
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_BLUE, BACKGROUND_BLUE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    styles.put(BoardStyles.BOARD_DARK_MIX_2_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentBs = getHeaderStyle(FONT_WHITE, BACKGROUND_PURPLE_HEAD, BORDER_BLACK, none);
    currentBs.setBold(false);
    currentStyle.put("HEAD", currentBs);
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_GREEN, BACKGROUND_GREEN_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    styles.put(BoardStyles.BOARD_DARK_MIX_3_STYLE, currentStyle);

    currentStyle = new HashMap<String, TableStyle>();
    currentBs = getHeaderStyle(FONT_WHITE, BACKGROUND_ORANGE_HEAD, BORDER_BLACK, none);
    currentBs.setBold(false);
    currentStyle.put("HEAD", currentBs);
    currentStyle.put("BODY_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("BODY_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_EVEN", getBodyStyle(true, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    currentStyle.put("FOOT_ODD", getBodyStyle(false, BACKGROUND_TURQUOISE, BACKGROUND_TURQUOISE_MEDIUM, FONT_BLACK, BORDER_BLACK, none));
    styles.put(BoardStyles.BOARD_DARK_MIX_4_STYLE, currentStyle);
  }



  /**
   * Save all the configuration for the headers of board in a BoardStyle. We can use
   * later this object to set all the style option
   * 
   * @param fontColor the font color
   * @param fillColor the background color
   * @param borderColor the border color
   * @param borderStyle the configuration of the cell's border
   * @return a complete configuration for a cell
   */
  private static TableStyle getHeaderStyle(XSSFColor fontColor, XSSFColor fillColor, XSSFColor borderColor, BorderStyle borderStyle) {
    TableStyle header = new TableStyle();
    borderStyle.completeBorderInfo(header);
    header.setBold(true);
    header.setFontColor(fontColor);
    header.setFillColor(fillColor);
    header.setBorderColor(borderColor);
    return header;
  }

  /**
   * Save all the configuration for the body of board in a BoardStyle. We can use
   * later this object to set all the style option
   * 
   * @param isEven indicate if the cell is in a row even or odd
   * @param evenColor the cell color for an even row
   * @param oddColor the cell color for an odd row
   * @param fontColor the font color
   * @param borderColor the border color
   * @param borderStyle the configuration of the cell's border
   * @return a complete configuration for a cell
   */
  private static TableStyle getBodyStyle(boolean isEven, XSSFColor evenColor, XSSFColor oddColor, 
                                          XSSFColor fontColor, XSSFColor borderColor, BorderStyle borderStyle) {
    TableStyle body = new TableStyle();
    body.setAlignment(CellStyle.ALIGN_LEFT);
    borderStyle.completeBorderInfo(body);
    body.setFontColor(fontColor);
    if (isEven) {
      body.setFillColor(evenColor);
    } else {
      body.setFillColor(oddColor);
    }
    body.setBorderColor(borderColor);
    return body;
  }
}
