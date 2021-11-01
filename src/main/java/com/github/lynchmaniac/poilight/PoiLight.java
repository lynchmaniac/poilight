package com.github.lynchmaniac.poilight;

import com.github.lynchmaniac.poilight.enumerations.BoardStyles;
import com.github.lynchmaniac.poilight.models.ExcelCell;
import com.github.lynchmaniac.poilight.models.ExcelRow;
import com.github.lynchmaniac.poilight.models.Table;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;



/**
 * This class is the entry point of this library. It contains methods to generate 
 * Excel spreadsheets as well as methods to generate complete Excel files.
 * If the user does not specify a style in leaf cells then the default style is applied. 
 * Otherwise it can use the predefined styles from Excel 2016. If it specifies a 
 * style on cells, it has priority over the predefined styles.
 * 
 * @author vpiard
 * @since 0.1
 */
public final class PoiLight {

  
  /**
   * Constructor.
   */
  public PoiLight() {
    
  }
  
  /**
   * Generate an Excel file.
   * 
   * @param filePath the full path where the file should be saved
   * @param data the data to be written in the spreadsheet
   */
  public static void generateExcel(String filePath, Table data) {
    XSSFWorkbook wb = new XSSFWorkbook();
    createTable(wb, data, false);
    writeExcel(wb, filePath);
  }

  /**
   * Generate an Excel file in streaming mode. Use for large file.
   * 
   * @param filePath the full path where the file should be saved
   * @param data the data to be written in the spreadsheet
   */
  public static void generateStreamingExcel(String filePath, Table data) {
    SXSSFWorkbook wb = new SXSSFWorkbook();
    createTable(wb, data, true);
    writeExcel(wb, filePath);
  }

  /**
   * Write a workbook in a file.
   * 
   * @param wb Excel Workbook
   * @param filePath the full path where the file should be saved
   */
  public static void writeExcel(Workbook wb, String filePath) {
    FileOutputStream fileOut = null;
    try {
      fileOut = new FileOutputStream(filePath);
      wb.write(fileOut);
    } catch (IOException exception) {
      System.out.println(exception);
    } finally {
      try {
        fileOut.close();
      } catch (IOException exception) {
        System.out.println(exception);
      }
    }
  }


  /**
   * Generate an SXSSF Excel's spreadsheet.
   * 
   * @param wb Excel Workbook
   * @param data the data to be written in the spreadsheet in format Table
   */
  public static void createStreamingTable(Workbook wb, Table data) {
    createTable(wb, data, true);
  }

  /**
   * Generate an XSSF Excel's spreadsheet.
   * 
   * @param wb Excel Workbook
   * @param data the data to be written in the spreadsheet in format Table
   */
  public static void createTable(Workbook wb, Table data) {
    createTable(wb, data, false);
  }

  /**
   * Generate an Excel spreadsheet.
   * 
   * @param wb Excel Workbook
   * @param data the data to be written in the spreadsheet
   * @param isStreaming indicates whether to write the streaming file
   * @param sheetName the name of the spreadsheet
   * @param bs all properties suitable on the style of a cell
   * @param firstRow the leading index of the table including the headers
   * @param firstCol the index of the first column
   */
  private static void createTable(Workbook wb, Table data, boolean isStreaming) {
    
    // determine the current sheet
    if (data.getSheetName() == null || "".equals(data.getSheetName())) {
      data.setSheetName(PoiLightConst.DEFAULT_SHEET_NAME);
    }

    Sheet sheet = wb.getSheet(data.getSheetName());
    if (sheet == null) {
      sheet = isStreaming ? (SXSSFSheet) wb.createSheet(data.getSheetName()) : (XSSFSheet) wb.createSheet(data.getSheetName());
    }

    // Control of position index
    data.setRow(controlRowIndex(wb, data.getData(), data.getRow()));
    data.setCol(controlColIndex(wb, data.getData(), data.getCol()));

    // Creation of style
    if (data.getStyle() == null) {
      data.setStyle(BoardStyles.BOARD_DEFAULT_STYLE);
    }
    // Row and column indexes
    int idx = data.getRow();
    int total = data.getData().size();
    int idxEven = 1;

    // Generate header
    if (data.hasHeader()) {
      Row row = sheet.getRow(idx);
      if (row == null) {
        row = sheet.createRow(idx);
      }
      int numberCol = data.getCol();
      for (ExcelCell cell : data.getHeader()) {
        Cell currentCell = row.getCell(numberCol);
        if (currentCell == null) {
          currentCell = row.createCell(numberCol);
        }
        setCellValue(currentCell, cell);
        applyCellStyle(wb, cell, currentCell, data.getStyle(), true, false, true);
        numberCol++;
      }
      idx++;
    }

    // Generate column headings
    for (ExcelRow currentRowContent : data.getData()) {

      Row row = sheet.getRow(idx);
      if (row == null) {
        row = sheet.createRow(idx);
      }
      int numberCol = data.getCol();
      for (ExcelCell cell : currentRowContent.getValue()) {
        Cell currentCell = row.getCell(numberCol);
        if (currentCell == null) {
          currentCell = row.createCell(numberCol);
        }
        setCellValue(currentCell, cell);
        boolean isFooter = total == idxEven;
        boolean isEven = !(idxEven % 2 == 0);
        applyCellStyle(wb, cell, currentCell, data.getStyle(), false, isFooter, isEven);
        numberCol++;
      } 
      idxEven++;
      idx++;
    }
  }

  /**
   * check the index of the first row. If the index is not included in the possible limits of Excel, 
   * it is then corrected to match the closest.
   * HSSF a type of workbook can only contain 65,536 rows.
   * XSSF a type of workbook or SXSSF only contain rows 1048576.
   * If the index is less than 1 then it is reduced to 1. If it is greater than the maximum terminal 
   * then returns the max index along the length of the table.
   * 
   * @param wb Excel Workbook
   * @param data the data to be written in the spreadsheet
   * @param firstRow the leading index of the table including the headers
   * @return controlled the index
   */
  private static Integer controlRowIndex(Workbook wb, List<ExcelRow> data, Integer firstRow) {
    Integer result = firstRow;
    Integer limit = firstRow + data.size();
    if (firstRow < 0) {
      result = 0;
    } else if (wb instanceof HSSFWorkbook && limit >= PoiLightConst.EXCEL_1997_MAX_ROW) {
      result = PoiLightConst.EXCEL_1997_MAX_ROW - data.size();
    } else if ((wb instanceof XSSFWorkbook || wb instanceof SXSSFWorkbook) && limit >= PoiLightConst.EXCEL_2007_MAX_ROW) {
      result = PoiLightConst.EXCEL_2007_MAX_ROW - data.size();
    }
    return result;
  }

  /**
   * check the index of the first column. If the index is not included in the possible limits of Excel,
   *  it is then corrected to match the closest.
   *  HSSF a type of workbook can contain only 256 columns.
   *  XSSF a type of workbook or SXSSF contain only 16385 columns.
   *  If the index is less than 1 then it is reduced to 1. If it is greater than the maximum terminal 
   *  then returns the index max following the width of the table.
   *  
   * @param wb Excel Workbook
   * @param data the data to be written in the spreadsheet
   * @param firstCol the index of the first column
   * @return controlled the index
   */
  private static Integer controlColIndex(Workbook wb, List<ExcelRow> data, Integer firstCol) {
    Integer result = firstCol;
    Integer limit = firstCol + data.get(0).getValue().size();
    if (firstCol < 0) {
      result = 0;
    } else if (wb instanceof HSSFWorkbook && limit >= PoiLightConst.EXCEL_1997_MAX_COL) {
      result = PoiLightConst.EXCEL_1997_MAX_COL - data.get(0).getValue().size();
    } else if ((wb instanceof XSSFWorkbook || wb instanceof SXSSFWorkbook) && limit >= PoiLightConst.EXCEL_2007_MAX_COL) {
      result = PoiLightConst.EXCEL_2007_MAX_COL - data.get(0).getValue().size();
    }
    return result;
  }

  /**
   * Applies the style to the cell. If the Cell Content object contains a 
   * style then it is applied, it takes precedence over the rest. Otherwise 
   * a predefined style is applied.
   * 
   * @param wb Excel Workbook
   * @param cell the content of the cell
   * @param c the cell
   * @param style all properties suitable on the style of a cell
   * @param isHeader indicates whether it is a header
   * @param isFooter indicates whether it is a footer
   * @param isEven indicates whether the current cell is on a par or odd row
   */
  private static void applyCellStyle(Workbook wb, ExcelCell cell, Cell currentCell, 
                                  BoardStyles style, boolean isHeader, boolean isFooter, boolean isEven) {
    if (cell.getStyle() != null) {
      currentCell.setCellStyle(cell.getStyle());
    } else {
      // This is a predefined style
      if (isHeader) {
        // apply the header style
        currentCell.setCellStyle(PoiLightStyle.getHeaderStyle(wb, style));
      } else if (isFooter) {
        // apply the footer style
        currentCell.setCellStyle(PoiLightStyle.getFooterStyle(wb, style, isEven));
      } else {
        // apply the body style
        currentCell.setCellStyle(PoiLightStyle.getBodyStyle(wb, style, isEven));
      }
    }
  }

  /**
   * Fill the cell value.
   * 
   * @param currentCell the cell
   * @param cell the content of the cell
   */
  private static void setCellValue(Cell currentCell, ExcelCell cell) {
    if (cell.getValue() instanceof String) {
      if (cell.isFormula()) {
        currentCell.setCellFormula((String) cell.getValue());
      } else {
        currentCell.setCellValue((String) cell.getValue());
      }
    }
    if (cell.getValue() instanceof Integer) {
      currentCell.setCellValue((Integer) cell.getValue());
    }
    if (cell.getValue() instanceof Long) {
      currentCell.setCellValue((Long) cell.getValue());
    }
    if (cell.getValue() instanceof Double) {
      currentCell.setCellValue((Double) cell.getValue());
    }
  }

}
