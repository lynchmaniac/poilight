package com.github.lynchmaniac.poilight;

import com.github.lynchmaniac.poilight.enumerations.BoardStyles;
import com.github.lynchmaniac.poilight.models.ExcelCell;
import com.github.lynchmaniac.poilight.models.Table;

import java.util.List;

/**
 * This class is a helper to manipulate the data.
 * 
 * @author vpiard
 * @since 0.1.1
 */
public class PoiLightData {

  /**
   * Return a table fully configured with all informations.
   * 
   * @param bs the predefined style
   * @return a table fully configure
   */
  public static Table getTable(BoardStyles bs) {
    return getTable(PoiLightConst.DEFAULT_SHEET_NAME, bs, PoiLightConst.DEFAULT_POSITION);
  }

  /**
   * Return a table fully configured with all informations.
   * 
   * @param sheetName the name of the current spreadsheet
   * @param bs the predefined style
   * @return a table fully configure
   */
  public static Table getTable(String sheetName, BoardStyles bs) {
    return getTable(sheetName, bs, PoiLightConst.DEFAULT_POSITION);
  }

  /**
   * Return a table fully configured with all informations.
   * 
   * @param sheetName the name of the current spreadsheet
   * @param bs the predefined style
   * @param position the current position for the table
   * @return a table fully configure
   */
  public static Table getTable(String sheetName, BoardStyles bs, String position) {
    Table table = new Table();
    table.setSheetName(sheetName);
    table.setStyle(bs);
    table.setPosition(position);
    return table;
  }


  /**
   * Return a table fully configured with all informations
   * about headers.
   * 
   * @param headers all the header of the table
   * @return a table fully configure
   */
  public static Table getTable(List<String> headers) {
    Table table = new Table();
    for (String header : headers) {
      table.addHeader(new ExcelCell(header));
    }
    return table;
  }


}
