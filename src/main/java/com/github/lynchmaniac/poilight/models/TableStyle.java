package com.github.lynchmaniac.poilight.models;

import com.github.lynchmaniac.poilight.PoiLightConst;
import com.github.lynchmaniac.poilight.helpers.StyleHelper;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * This object is used to store the many configuration of the Excel's styles.
 * This object properies is after apply to the real properties of POI.
 * 
 * @author vpiard
 * @since 0.1
 */
public class TableStyle {

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
  private Short fontSize = new Short("11");
  /**
   * Indicate the cell's alignment.
   */
  private Short alignment = CellStyle.ALIGN_LEFT;

  /**
   * The style of the border left of the cell.
   */
  private Short cellBorderLeft = CellStyle.BORDER_MEDIUM;

  /**
   * The style of the border right of the cell.
   */
  private Short cellBorderRight = CellStyle.BORDER_MEDIUM;

  /**
   * The style of the border top of the cell.
   */
  private Short cellBorderTop = CellStyle.BORDER_THIN;

  /**
   * The style of the border bottom of the cell.
   */
  private Short cellBorderBottom = CellStyle.BORDER_THIN;


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
      boolean isBold, Short fontSize, Short alignment) {
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

  public Short getFontSize() {
    return fontSize;
  }

  public void setFontSize(Short fontSize) {
    this.fontSize = fontSize;
  }

  public Short getAlignment() {
    return alignment;
  }

  public void setAlignment(Short alignment) {
    this.alignment = alignment;
  }

  public Short getCellBorderLeft() {
    return cellBorderLeft;
  }

  public void setCellBorderLeft(Short cellBorderLeft) {
    this.cellBorderLeft = cellBorderLeft;
  }

  public Short getCellBorderRight() {
    return cellBorderRight;
  }

  public void setCellBorderRight(Short cellBorderRight) {
    this.cellBorderRight = cellBorderRight;
  }

  public Short getCellBorderTop() {
    return cellBorderTop;
  }

  public void setCellBorderTop(Short cellBorderTop) {
    this.cellBorderTop = cellBorderTop;
  }

  public Short getCellBorderBottom() {
    return cellBorderBottom;
  }

  public void setCellBorderBottom(Short cellBorderBottom) {
    this.cellBorderBottom = cellBorderBottom;
  }

}
