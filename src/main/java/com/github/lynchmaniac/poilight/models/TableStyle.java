package com.github.lynchmaniac.poilight.models;

import com.github.lynchmaniac.poilight.PoiLightConst;
import com.github.lynchmaniac.poilight.helpers.StyleHelper;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.Serializable;

/**
 * This object is used to store the many configuration of the Excel's styles.
 * This object properies is after apply to the real properties of POI.
 * 
 * @author vpiard
 * @since 0.1
 */
public class TableStyle implements Serializable {

  /**
   * UID.
   */
  private static final long serialVersionUID = -6584637308090528538L;
  /**
   * The font color.
   */
  private XSSFColor fontColor = StyleHelper.getColor(0, 0, 0);
  /**
   * The color of the border of the cell.
   */
  private XSSFColor borderColor = StyleHelper.getColor(0, 0, 0);
  /**
   * The style of the background of the cell.
   */
  private XSSFColor fillColor = null;
  /**
   * The font name.
   */
  private String fontName = PoiLightConst.EXCEL_2007_DEFAULT_FONT_NAME;
  /**
   * Indicate if the font is Bold.
   */
  private boolean isBold = false;
  /**
   * The font size.
   */
  private short fontSize = 11;
  /**
   * Indicate the cell's alignment.
   */
  private short alignment = CellStyle.ALIGN_LEFT;

  /**
   * The style of the border left of the cell.
   */
  private short cellBorderLeft = CellStyle.BORDER_MEDIUM;

  /**
   * The style of the border right of the cell.
   */
  private short cellBorderRight = CellStyle.BORDER_MEDIUM;

  /**
   * The style of the border top of the cell.
   */
  private short cellBorderTop = CellStyle.BORDER_THIN;

  /**
   * The style of the border bottom of the cell.
   */
  private short cellBorderBottom = CellStyle.BORDER_THIN;


  /**
   * Constructor.
   */
  public TableStyle() {
    super();
  }

  /**
   * Constructor.
   * 
   * @param fillColor the background color
   */
  public TableStyle(XSSFColor fillColor) {
    super();
    this.fillColor = fillColor;
  }


  /**
   * Constructor.
   * 
   * @param fontColor the font color
   * @param fillColor the background color
   * @param fontName the font name
   * @param isBold indicate if the font is bold
   * @param fontSize the font size
   * @param alignment the cell's alignment
   */
  public TableStyle(XSSFColor fontColor, XSSFColor fillColor, String fontName,
      boolean isBold, short fontSize, short alignment) {
    super();
    this.fontColor = fontColor;
    this.fillColor = fillColor;
    this.fontName = fontName;
    this.isBold = isBold;
    this.fontSize = fontSize;
    this.alignment = alignment;
  }


  public XSSFColor getFontColor() {
    return fontColor;
  }

  public void setFontColor(XSSFColor fontColor) {
    this.fontColor = fontColor;
  }

  public XSSFColor getFillColor() {
    return fillColor;
  }

  public XSSFColor getBorderColor() {
    return borderColor;
  }

  public void setBorderColor(XSSFColor borderColor) {
    this.borderColor = borderColor;
  }

  public void setFillColor(XSSFColor fillColor) {
    this.fillColor = fillColor;
  }

  public String getFontName() {
    return fontName;
  }

  public void setFontName(String fontName) {
    this.fontName = fontName;
  }

  public boolean isBold() {
    return isBold;
  }

  public void setBold(boolean isBold) {
    this.isBold = isBold;
  }

  public short getFontSize() {
    return fontSize;
  }

  public void setFontSize(short fontSize) {
    this.fontSize = fontSize;
  }

  public short getAlignment() {
    return alignment;
  }

  public void setAlignment(short alignment) {
    this.alignment = alignment;
  }

  public short getCellBorderLeft() {
    return cellBorderLeft;
  }

  public void setCellBorderLeft(short cellBorderLeft) {
    this.cellBorderLeft = cellBorderLeft;
  }

  public short getCellBorderRight() {
    return cellBorderRight;
  }

  public void setCellBorderRight(short cellBorderRight) {
    this.cellBorderRight = cellBorderRight;
  }

  public short getCellBorderTop() {
    return cellBorderTop;
  }

  public void setCellBorderTop(short cellBorderTop) {
    this.cellBorderTop = cellBorderTop;
  }

  public short getCellBorderBottom() {
    return cellBorderBottom;
  }

  public void setCellBorderBottom(short cellBorderBottom) {
    this.cellBorderBottom = cellBorderBottom;
  }

}
