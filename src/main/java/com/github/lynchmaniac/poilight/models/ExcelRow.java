package com.github.lynchmaniac.poilight.models;

import java.util.ArrayList;
import java.util.List;



/**
 * A description of the contents of a cell
 * 
 * @author vpiard
 * @since 0.1
 */
public class ExcelRow {

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
    if (value == null) {
      value = new ArrayList<ExcelCell>();
    }
    return value;
  }

  /**
   * Add a ExcelCell of the current row.
   * 
   * @param value a ExcelCell
   */
  public void addValue(ExcelCell value) {
    if (this.value == null) {
      this.value = new ArrayList<ExcelCell>();
    }
    this.value.add(value);
  }
}
