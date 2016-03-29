package com.github.lynchmaniac.poilight.models;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;



/**
 * A description of the contents of a cell
 * 
 * @author vpiard
 * @since 0.1
 */
public class ExcelRow implements Serializable {

  /**
   * UID.
   */
  private static final long serialVersionUID = 6578060542569742883L;
  
  /**
   * All the cell of a row.
   */
  private List<ExcelCell> value;

  /**
   * Constructor.
   */
  public ExcelRow() {
    super();
  }

  /**
   * Constructor.
   * You can put all the ExcelCell in parameters you want.
   * It's useful when you want to create a row easily.
   *
   * @param datas a number undefined of ExcellCell
   */
  public ExcelRow(ExcelCell ... datas) {
    super();
    for (ExcelCell cellContent : datas) {
      addValue(cellContent);
    }
  }

  /**
   * Constructor.
   * You can put all the Object in parameters you want.
   * It's useful when you want to create a row easily.
   * But with this method you can't style on cell.
   *
   * @param datas a number undefined of ExcellCell
   */
  public ExcelRow(Object ... datas) {
    super();
    for (Object content : datas) {
      addValue(new ExcelCell(content));
    }
  }
  
  public ExcelRow(List<ExcelCell> value) {
    super();
    this.value = value;
  }

  /**
   * Returns all the ExcelCell of the row.
   * 
   * @return all the ExcelCell of the row
   */
  public List<ExcelCell> getValue() {
    initializeList();
    return value;
  }

  /**
   * Add a ExcelCell of the current row.
   * 
   * @param value a ExcelCell
   */
  public void addValue(ExcelCell value) {
    initializeList();
    this.value.add(value);
  }
  
  /**
   * Check if the list is null and instanciate a new if it's true.
   */
  private void initializeList() {
    if (this.value == null) {
      this.value = new ArrayList<ExcelCell>();
    }
  }
}
