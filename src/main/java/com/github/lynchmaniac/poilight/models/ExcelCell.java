package com.github.lynchmaniac.poilight.models;

import org.apache.poi.ss.usermodel.CellStyle;

import java.io.Serializable;



/**
 * This object resume the content of a cell.
 * You can customize the style of the cell with the property style.
 * If you want to highlight the cell, you must set the property color to true.
 * If you want set a formula in the cell, you must set the property formula to true.
 * 
 * @author vpiard
 * @since 0.1
 */
public class ExcelCell implements Serializable {

  /**
   * UID.
   */
  private static final long serialVersionUID = -3458532388085217659L;
  
  /**
   * Represent the content of the cell.
   * It can be whatever you want : String, numerical, date...
   */
  private transient Object value;
  /**
   * Indicates if the cell must be colored or not.
   * It's useful when you want to highlight a cell.
   */
  private boolean color;
  /**
   * Represent the style of the cell.
   */
  private transient CellStyle style;
  
  /**
   * Indicate if the content of the cell is a formula.
   */
  private boolean formula;

  /**
   * Constructor.
   * 
   * @param value the cell's value
   */
  public ExcelCell(Object value) {
    super();
    this.value = value;
    this.color = false;
  }
  
  /**
   * Constructor.
   * 
   * @param value the cell's value
   * @param style the cell's style
   */
  public ExcelCell(Object value, CellStyle style) {
    super();
    this.value = value;
    this.style = style;
  }

  /**
   * Constructor.
   * 
   * @param value the cell's value
   * @param isFormula indicate if the content of the cell is a formula
   */
  public ExcelCell(Object value, boolean isFormula) {
    super();
    this.value = value;
    this.formula = isFormula;
  }


  /**
   * Constructor.
   * 
   * @param value the cell's value
   * @param isFormula indicate if the content of the cell is a formula
   * @param style the cell's style
   */
  public ExcelCell(Object value, boolean isFormula, CellStyle style) {
    super();
    this.value = value;
    this.formula = isFormula;
    this.style = style;
  }

  /**
   * Constructor.
   * 
   * @param value the cell's value
   * @param isFormula indicate if the content of the cell is a formula
   * @param style the cell's style
   * @param color indicates whether the cell should be colored
   */
  public ExcelCell(Object value, boolean isFormula, CellStyle style, boolean color) {
    super();
    this.value = value;
    this.color = color;
    this.style = style;
    this.formula = isFormula;
  }

  public Object getValue() {
    return value;
  }

  public void setValue(String value) {
    this.value = value;
  }

  public void setValue(Double value) {
    this.value = value;
  }

  public void setValue(Integer value) {
    this.value = value;
  }

  public void setValue(Long value) {
    this.value = value;
  }

  public boolean isColor() {
    return color;
  }

  public void setColor(boolean color) {
    this.color = color;
  }

  public CellStyle getStyle() {
    return style;
  }

  public void setStyle(CellStyle style) {
    this.style = style;
  }

  public boolean isFormula() {
    return formula;
  }

  public void setFormula(boolean formula) {
    this.formula = formula;
  }

}
