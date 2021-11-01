package com.github.lynchmaniac.poilight;

import com.github.lynchmaniac.poilight.enumerations.BoardStyles;
import com.github.lynchmaniac.poilight.helpers.CreateExcelStyleHelper;
import com.github.lynchmaniac.poilight.models.TableStyle;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;


/**
 * This class contains all the methods to apply the style to the cells of the Excel workbook
 * 
 * @author vpiard
 * @since 0.1
 */
public final class PoiLightStyle {

  /**
   * Constructor.
   */
  private PoiLightStyle() {
    throw new IllegalStateException("Utility class");
  }
  
  /**
   * Create and send a form from the saved configuration.
   * 
   * @param workbook Excel Workbook
   * @param boardStyle all properties suitable on the style of a cell
   * @return a font
   */
  protected static Font getFontStyle(Workbook workbook, TableStyle boardStyle) {
    XSSFFont font = (XSSFFont) workbook.createFont();
    font.setFontName(boardStyle.getFontName());
    font.setFontHeightInPoints(boardStyle.getFontSize());
    font.setBold(boardStyle.isBold());
    font.setColor(boardStyle.getFontColor());
    return font;
  }

  /**
   * Fixed all properties suitable for cell-related style.
   * 
   * @param workbook Excel Workbook
   * @param boardStyle all properties suitable on the style of a cell
   * @param font a font
   * @return the customized style
   */
  protected static CellStyle getCellStyle(Workbook workbook, TableStyle boardStyle, Font font) {
    XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
    if (boardStyle.getFillColor() != null) {
      cellStyle.setFillForegroundColor(boardStyle.getFillColor());
      cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
    cellStyle.setBorderLeft(boardStyle.getCellBorderLeft());
    cellStyle.setBorderRight(boardStyle.getCellBorderRight());
    cellStyle.setBorderTop(boardStyle.getCellBorderTop());
    cellStyle.setBorderBottom(boardStyle.getCellBorderBottom());
    cellStyle.setAlignment(boardStyle.getAlignment());

    cellStyle.setBorderColor(BorderSide.LEFT, boardStyle.getBorderColor());
    cellStyle.setBorderColor(BorderSide.RIGHT, boardStyle.getBorderColor());
    cellStyle.setBorderColor(BorderSide.TOP, boardStyle.getBorderColor());
    cellStyle.setBorderColor(BorderSide.BOTTOM, boardStyle.getBorderColor());

    if (font != null) {
      cellStyle.setFont(font);
    }
    return cellStyle;
  }


  /**
   * returns a cell style complemented with all previously-saved Details.
   * 
   * @param wb Excel Workbook
   * @param boardStyle all properties suitable on the style of a cell
   * @return a cell style 
   */
  protected static CellStyle createStyle(Workbook wb, TableStyle boardStyle) {
    Font font = getFontStyle(wb, boardStyle);
    CellStyle cs = getCellStyle(wb, boardStyle, font);
    return cs;
  }

  /**
   * Returns the style of the cell to the headers.
   * 
   * @param wb Excel Workbook
   * @param style all properties suitable on the style of a cell
   * @return the style of a header
   */
  protected static CellStyle getHeaderStyle(Workbook wb, BoardStyles style) {
    return createStyle(wb , CreateExcelStyleHelper.getExcelStyle().get(style).get("HEAD"));
  }
  
  /**
   * Returns the style of the cell to the body cell.
   * 
   * @param wb Excel Workbook
   * @param style all properties suitable on the style of a cell
   * @param isEven indicates whether the current cell is on a par or odd row
   * @return the style of a body cell 
   */
  protected static CellStyle getBodyStyle(Workbook wb, BoardStyles style, boolean isEven) {
    return getStyle(wb, style, isEven, true);
  }

  /**
   * Returns the style of the cell footers.
   * 
   * @param wb Excel Workbook
   * @param style all properties suitable on the style of a cell
   * @param isEven indicates whether the current cell is on a par or odd row
   * @return style footer
   */
  protected static CellStyle getFooterStyle(Workbook wb, BoardStyles style, boolean isEven) {
    return getStyle(wb, style, isEven, false);
  }

  /**
   * Generates the style for all cells except for headers.
   * 
   * @param wb Excel Workbook
   * @param style all properties suitable on the style of a cell
   * @param isEven indicates whether the current cell is on a par or odd row
   * @param isBody indicates whether the cell is part of the body and not footers
   * @return the cell style for all the cell except the headers
   */
  private static CellStyle getStyle(Workbook wb, BoardStyles style, boolean isEven, boolean isBody) {
    CellStyle result;
    String keyEven = isBody ? "BODY_EVEN" : "FOOT_EVEN";
    String keyOdd = isBody ? "BODY_ODD" : "FOOT_ODD";
    if (isEven || CreateExcelStyleHelper.getExcelStyle().get(style).get(keyOdd) == null) {
      result = createStyle(wb , CreateExcelStyleHelper.getExcelStyle().get(style).get(keyEven));
    } else {
      result = createStyle(wb , CreateExcelStyleHelper.getExcelStyle().get(style).get(keyOdd));
    }
    return result;
  }
}
