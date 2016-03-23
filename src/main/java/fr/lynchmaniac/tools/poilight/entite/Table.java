package fr.lynchmaniac.tools.poilight.entite;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import fr.lynchmaniac.tools.poilight.enumeration.BoardStyles;

public class Table {

	private List<RowContent> data;
	private List<CellContent> header;
	private String sheetName = "data";
	private Integer row = 0;
	private Integer col = 0;
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

	public void setPosition(String position) {
		if (position == null || "".equals(position)) {
			row = 0;
			col = 0;
		}
		Pattern pattern = Pattern.compile("^([A-Za-z]{1,3})([1-9]{1,})");
		Matcher matcher = pattern.matcher(position);
		boolean b = matcher.matches();
		// si recherche fructueuse
		if(b) {
			col = getColNum(matcher.group(1)) - 1;
			row = Integer.valueOf(matcher.group(2)) - 1;
		} else {
			row = 0;
			col = 0;
		}

	}

	private int getColNum (String colName) {

		//remove any whitespace
		colName = colName.trim();

		StringBuffer buff = new StringBuffer(colName);

		//string to lower case, reverse then place in char array
		char chars[] = buff.reverse().toString().toLowerCase().toCharArray();

		int retVal=0, multiplier=0;

		for(int i = 0; i < chars.length;i++){
			//retrieve ascii value of character, subtract 96 so number corresponds to place in alphabet. ascii 'a' = 97 
			multiplier = (int)chars[i]-96;
			//mult the number by 26^(position in array)
			retVal += multiplier * Math.pow(26, i);
		}
		return retVal;
	}

	//	public String getPosition() {
	//		// FIXME calcul complexe à faire pour générer la position en fonction des row & col
	//		return "A1";
	//	}

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
