package com.github.lynchmaniac.poilight.entity;

import java.util.ArrayList;
import java.util.List;



/**
 * A description of the contents of a cell
 * 
 * @author vpiard
 * @since 0.1
 */
public class RowContent {

	/**
	 * All the cell of a row.
	 */
	private List<CellContent> value;

	/**
	 * Constructor.
	 */
	public RowContent() {
		super();
	}
	
	public RowContent(CellContent ... datas) {
		super();
		for (CellContent cellContent : datas) {
			addValue(cellContent);
		}
	}

	/**
	 * @param value the cell value
	 */
	public RowContent(List<CellContent> value) {
		super();
		this.value = value;
	}

	/**
	 * @return value the cell value
	 */
	public List<CellContent> getValue() {
		if (value == null) {
			value = new ArrayList<CellContent>();
		}
		return value;
	}

	/**
	 * @param value value to define
	 */
	public void addValue(CellContent value) {
		if (this.value == null) {
			this.value = new ArrayList<CellContent>();
		}
		this.value.add(value);
	}
}
