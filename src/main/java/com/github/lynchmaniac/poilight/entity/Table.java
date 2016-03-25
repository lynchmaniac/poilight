package com.github.lynchmaniac.poilight.entity;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.github.lynchmaniac.poilight.enumeration.BoardStyles;

/**
 * This class contain all the data to draw a table.
 * 
 * @author vpiard
 * @since 0.1
 */
public class Table {

	private static Pattern pattern = Pattern.compile("^([A-Za-z]{1,3})([1-9]{1,})");
	/**
	 * The data for the spreadsheet.
	 * You can add style for each data.
	 */
	private List<RowContent> data;
	/**
	 * The header content of the table.
	 * You can add style for each hearder.
	 */
	private List<CellContent> header;
	/**
	 * The nam of the spreadsheet
	 */
	private String sheetName = "data";
	/**
	 * The row of the first table cell, the top left.
	 */
	private Integer row = 0;
	/**
	 * The col of the first table cell, the top left.
	 */
	private Integer col = 0;
	/**
	 * The predifened Excel's style
	 */
	private BoardStyles style = BoardStyles.BOARD_DEFAULT_STYLE;

	/**
	 * @return the data
	 */
	public List<RowContent> getData() {
		if (data == null) {
			data = new ArrayList<RowContent>();
		}
		return data;
	}

	/**
	 * @param data
	 *            the data to set
	 */
	public void setData(List<RowContent> data) {
		this.data = data;
	}
	
	public void addData(RowContent data) {
		if (this.data == null) {
			this.data = new ArrayList<RowContent>();
		}
		this.data.add(data);
	}

	/**
	 * @return the header
	 */
	public List<CellContent> getHeader() {
		if (this.header == null) {
			this.header = new ArrayList<CellContent>();
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
	 * @param header
	 *            the header to set
	 */
	public void setHeader(List<CellContent> header) {
		this.header = header;
	}
	
	public void addHeader(CellContent header) {
		if (this.header == null) {
			this.header = new ArrayList<CellContent>();
		}
		this.header.add(header);
	}

	/**
	 * @return the sheetName
	 */
	public String getSheetName() {
		return sheetName;
	}

	/**
	 * @param sheetName
	 *            the sheetName to set
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
		boolean b = matcher.matches();
		// If the regex matche
		if(b) {
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
	 * @return
	 */
	private int getColNum(String colName) {

		
		StringBuffer buff = new StringBuffer(colName.trim());
		//string to lower case, reverse then place in char array
		char chars[] = buff.reverse().toString().toLowerCase().toCharArray();

		int retVal=0, multiplier=0;
		for(int i = 0; i < chars.length;i++){
			//get ascii value
			multiplier = (int)chars[i]-96;
			//check for position
			retVal += multiplier * Math.pow(26, i);
		}
		return retVal;
	}

	/**
	 * @return the row
	 */
	public Integer getRow() {
		return row;
	}

	/**
	 * @param row
	 *            the row to set
	 */
	public void setRow(Integer row) {
		this.row = row;
	}

	/**
	 * @return the col
	 */
	public Integer getCol() {
		return col;
	}

	/**
	 * @param col
	 *            the col to set
	 */
	public void setCol(Integer col) {
		this.col = col;
	}

	/**
	 * @return the style
	 */
	public BoardStyles getStyle() {
		return style;
	}

	/**
	 * @param style the style to set
	 */
	public void setStyle(BoardStyles style) {
		this.style = style;
	}

}
