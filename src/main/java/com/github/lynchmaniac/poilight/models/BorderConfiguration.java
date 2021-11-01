package com.github.lynchmaniac.poilight.models;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Configure all the border style for each side of the cell.
 * You can have one specific style for every side of the cell.
 * 
 * @author vpiard
 * @since 0.1
 */
public class BorderConfiguration {

  /**
   * The bottom cell style.
   */
  private BorderStyle bottom = BorderStyle.NONE;
  /**
   * The top cell style.
   */
  private BorderStyle top = BorderStyle.NONE;
  /**
   * The left cell style.
   */
  private BorderStyle left = BorderStyle.NONE;
  /**
   * The right cell style.
   */
  private BorderStyle right = BorderStyle.NONE;

  /**
   * Constructor.
   */
  public BorderConfiguration() {
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
  public BorderConfiguration(BorderStyle bottom, BorderStyle top, BorderStyle left, BorderStyle right) {
    super();
    this.bottom = bottom;
    this.top = top;
    this.left = left;
    this.right = right;
  }
  
  public BorderStyle getBottom() {
    return bottom;
  }
  
  public void setBottom(BorderStyle bottom) {
    this.bottom = bottom;
  }
  
  public BorderStyle getTop() {
    return top;
  }
  
  public void setTop(BorderStyle top) {
    this.top = top;
  }
  
  public BorderStyle getLeft() {
    return left;
  }
  
  public void setLeft(BorderStyle left) {
    this.left = left;
  }
  
  public BorderStyle getRight() {
    return right;
  }
  
  public void setRight(BorderStyle right) {
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
