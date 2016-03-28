package com.github.lynchmaniac.poilight.models;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Configure all the border style for each side of the cell.
 * You can have one specific style for every side of the cell.
 * 
 * @author vpiard
 * @since 0.1
 */
public class BorderStyle {

  /**
   * The bottom cell style.
   */
  private short bottom = CellStyle.BORDER_NONE;
  /**
   * The top cell style.
   */
  private short top = CellStyle.BORDER_NONE;
  /**
   * The left cell style.
   */
  private short left = CellStyle.BORDER_NONE;
  /**
   * The right cell style.
   */
  private short right = CellStyle.BORDER_NONE;

  /**
   * Constructor.
   */
  public BorderStyle() {
    super();
  }

  /**
   * Constructor.
   * 
   * @param bottom the bottom cell style
   * @param top the top cell style
   * @param left the left cell style
   * @param right the right cell style
   */
  public BorderStyle(short bottom, short top, short left, short right) {
    super();
    this.bottom = bottom;
    this.top = top;
    this.left = left;
    this.right = right;
  }
  
  public short getBottom() {
    return bottom;
  }
  
  public void setBottom(short bottom) {
    this.bottom = bottom;
  }
  
  public short getTop() {
    return top;
  }
  
  public void setTop(short top) {
    this.top = top;
  }
  
  public short getLeft() {
    return left;
  }
  
  public void setLeft(short left) {
    this.left = left;
  }
  
  public short getRight() {
    return right;
  }
  
  public void setRight(short right) {
    this.right = right;
  }

  /**
   * Complete all fields concerning information on the cell borders
   * into the global object Table.
   *  
   * @param style the configuration's style
   */
  public void completeBorderInfo(TableStyle style) {
    style.setCellBorderBottom(getBottom());
    style.setCellBorderTop(getTop());
    style.setCellBorderLeft(getLeft());
    style.setCellBorderRight(getRight());
  }
}
