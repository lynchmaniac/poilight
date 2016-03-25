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
	
	public ExcelRow(ExcelCell ... datas) {
		super();
		for (ExcelCell cellContent : datas) {
			addValue(cellContent);
		}
	}

	/**
	 * @param value the cell value
	 */
	public ExcelRow(List<ExcelCell> value) {
		super();
		this.value = value;
	}

	/**
	 * @return value the cell value
	 */
	public List<ExcelCell> getValue() {
		if (value == null) {
			value = new ArrayList<ExcelCell>();
		}
		return value;
	}

	/**
	 * @param value value to define
	 */
	public void addValue(ExcelCell value) {
		if (this.value == null) {
			this.value = new ArrayList<ExcelCell>();
		}
		this.value.add(value);
	}
}
