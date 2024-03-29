package com.github.lynchmaniac.poilight;

/**
 * This class contains all constants related to different version of Excel
 * 
 * @author vpiard
 * @since 0.1
 */
public final class PoiLightConst {

  /**
   * Constructor.
   */
  private PoiLightConst() {
    throw new IllegalStateException("Const class");
  }
  
  /**
   * The default sheet name.
   */
  public static final String DEFAULT_SHEET_NAME = "data";

  /**
   * the default position of a table.
   */
  public static final String DEFAULT_POSITION = "A1";

  /**
   * In Excel 97, the maximum number of column.
   */
  public static final Integer EXCEL_1997_MAX_COL = 256;
  /**
   * In Excel 97, the maximum number of row.
   */
  public static final Integer EXCEL_1997_MAX_ROW = 65536;
  /**
   * In Excel 97, the maximum number of arguments passed to a function.
   */
  public static final Integer EXCEL_1997_MAX_ARGS = 30;
  /**
   * In Excel 97, the maximum number of style associated with a cell.
   */
  public static final Integer EXCEL_1997_NB_CELL_STYLES = 4000;
  /**
   * In Excel 97, the maximum size of text in a cell. 
   */
  public static final Integer EXCEL_1997_LENGTH_TEXT = 32767;
  /**
   * In Excel 97, the name of the default font.
   */
  public static final String EXCEL_1997_DEFAULT_FONT_NAME = "Arial";


  /**
   * In Excel 2007, the maximum number of column.
   */
  public static final Integer EXCEL_2007_MAX_COL = 16384;
  /**
   * In Excel 2007, the maximum number of row.
   */
  public static final Integer EXCEL_2007_MAX_ROW = 1048576;
  /**
   * In Excel 2007, the maximum number of arguments passed to a function.
   */
  public static final Integer EXCEL_2007_MAX_ARGS = 255;
  /**
   * In Excel 2007, the maximum number of style associated with a cell.
   */
  public static final Integer EXCEL_2007_NB_CELL_STYLES = 64000;
  /**
   * In Excel 2007, the maximum size of text in a cell. 
   */
  public static final Integer EXCEL_2007_LENGTH_TEXT = 32767;
  /**
   * In Excel 2007, the name of the default font.
   */
  public static final String EXCEL_2007_DEFAULT_FONT_NAME = "Calibri";

}
