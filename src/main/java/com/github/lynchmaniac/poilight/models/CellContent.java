package com.github.lynchmaniac.poilight.models;

import org.apache.poi.ss.usermodel.CellStyle;



/**
 * This object resume the content of a cell
 * 
 * @author vpiard
 * @since 0.1
 */
public class CellContent {

	private Object value;
	private boolean color;
	private CellStyle style;
	
	/**
	 * Constructor.
	 * 
	 * @param value the cell's value
	 */
	public CellContent(Object value) {
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
	public CellContent(Object value, CellStyle style) {
		super();
		this.value = value;
		this.style = style;
	}
	
	/**
	 * Constructor.
	 * 
	 * @param value the cell's value
	 * @param color indicates whether the cell should be colored
	 */
	public CellContent(Object value, boolean color) {
		super();
		this.value = value;
		this.color = color;
	}


	/**
	 * Constructor.
	 * 
	 * @param value the cell's value
	 * @param color indicates whether the cell should be colored
	 * @param style the cell's style
	 */
	public CellContent(Object value, boolean color, CellStyle style) {
		super();
		this.value = value;
		this.color = color;
		this.style = style;
	}
	/**
	 * @return value
	 */
	public Object getValue() {
		return value;
	}

	/**
	 * @param value value to define
	 */
	public void setValue(String value) {
		this.value = value;
	}
	
	/**
	 * @param value value to define
	 */
	public void setValue(Double value) {
		this.value = String.valueOf(value);
	}
	
	/**
	 * @param value value to define
	 */
	public void setValue(Integer value) {
		this.value = String.valueOf(value);
	}
	
	/**
	 * @param value value to define
	 */
	public void setValue(Long value) {
		this.value = String.valueOf(value);
	}

	/**
	 * @return color
	 */
	public boolean isColor() {
		return color;
	}

	/**
	 * @param color color to define
	 */
	public void setColor(boolean color) {
		this.color = color;
	}

	/**
	 * @return style
	 */
	public CellStyle getStyle() {
		return style;
	}

	/**
	 * @param style style to define
	 */
	public void setStyle(CellStyle style) {
		this.style = style;
	}
	
}
