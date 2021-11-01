package com.github.lynchmaniac.poilight.models;

import com.github.lynchmaniac.poilight.enumerations.BoardStyles;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * This class contains all the data to draw a table.
 * 
 * @author vpiard
 * @since 0.1
 */
public class Table implements Serializable {

  /**
   * UID.
   */
  private static final long serialVersionUID = 5597446760000118027L;

  private static final Pattern pattern = Pattern.compile("^([A-Za-z]{1,3})([1-9]{1,})");
  
  /**
   * The data for the spreadsheet.
   * You can add style for each data.
   */
  private List<ExcelRow> data;
  
  /**
   * The header content of the table.
   * You can add style for each hearder.
   */
  private List<ExcelCell> header;
 
  /**
   * The nam of the spreadsheet.
   */
  private String sheetName = "data";
  
  /**
   * The row of the first table cell, the top left.
   */
  private int row = 0;
  
  /**
   * The col of the first table cell, the top left.
   */
  private int col = 0;
  
  /**
   * The predifened Excel's style.
   */
  private BoardStyles style = BoardStyles.BOARD_DEFAULT_STYLE;

  /**
   * Return all the row of the table.
   * 
   * @return the data all the row of the table
   */
  public List<ExcelRow> getData() {
    if (data == null) {
      data = new ArrayList<ExcelRow>();
    }
    return data;
  }

  /**
   * Add a row to the current table.
   * 
   * @param data a row of the table
   */
  public void addData(ExcelRow data) {
    if (this.data == null) {
      this.data = new ArrayList<ExcelRow>();
    }
    this.data.add(data);
  }

  /**
   * Return the list of the cell which is the header Cell.
   * 
   * @return the header all the cell for the header
   */
  public List<ExcelCell> getHeader() {
    if (this.header == null) {
      this.header = new ArrayList<ExcelCell>();
    }
    return header;
  }
  
  /**
   * Indicate if the table has a header or not.
   * If header list is empty, so there are no header ;-)
   * 
   * @return true if the table has a header
   */
  public boolean hasHeader() {
    return !getHeader().isEmpty();
  }


  /**
   * Add a cell for the header.
   *  
   * @param header a cell for the header
   */
  public void addHeader(ExcelCell header) {
    if (this.header == null) {
      this.header = new ArrayList<ExcelCell>();
    }
    this.header.add(header);
  }
  
  /**
   * Fix all the ExcellCell header in one time.
   * 
   * @param datas all the header  
   */
  public void addHeaders(ExcelCell ... datas) {
    for (ExcelCell currentHeader : datas) {
      addHeader(currentHeader);
    }
  }
  
  /**
   * Fix all the object header in one time.
   * 
   * @param datas all the header  
   */
  public void addHeaders(Object ... datas) {
    for (Object content : datas) {
      addHeader(new ExcelCell(content));
    }
  }

  /**
   * Return the name of the sheet which must
   * contains the futur table.
   * 
   * @return the sheetName
   */
  public String getSheetName() {
    return sheetName;
  }

  /**
   * Fix the name of the sheet which must
   * contains the futur table.
   * 
   * @param sheetName the name of the current spreadsheet 
   */
  public void setSheetName(String sheetName) {
    this.sheetName = sheetName;
  }

  /**
   * This method is to specify the col and the row of the first 
   * table cell, the top left. You  use a Excel position like A1 or B78.
   * This method transform this notation in Integer for row and col.
   * 
   * @param position You must use a Excel position like A1 or B78
   */
  public void setPosition(String position) {
    if (position == null || "".equals(position)) {
      row = 0;
      col = 0;
    }

    Matcher matcher = pattern.matcher(position);
    // If the regex matche
    if (matcher.matches()) {
      col = getColNum(matcher.group(1)) - 1;
      row = Integer.valueOf(matcher.group(2)) - 1;
    } else {
      row = 0;
      col = 0;
    }

  }

  /**
   * Transform a column name into an Integer.
   * 
   * @param colName the name of a column
   * @return the number of the column
   */
  private int getColNum(String colName) {


    StringBuilder  buff = new StringBuilder(colName.trim());
    //string to lower case, reverse then place in char array
    char[] chars = buff.reverse().toString().toLowerCase().toCharArray();

    int retVal = 0;
    int multiplier = 0;
    for (int i = 0; i < chars.length;i++) {
      //get ascii value
      multiplier = chars[i] - 96;
      //check for position
      retVal += multiplier * Math.pow(26, i);
    }
    return retVal;
  }


  public Integer getRow() {
    return row;
  }

  public void setRow(Integer row) {
    this.row = row;
  }

  public Integer getCol() {
    return col;
  }

  public void setCol(Integer col) {
    this.col = col;
  }

  public BoardStyles getStyle() {
    return style;
  }

  public void setStyle(BoardStyles style) {
    this.style = style;
  }

}
