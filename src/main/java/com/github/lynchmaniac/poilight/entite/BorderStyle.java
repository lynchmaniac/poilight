package com.github.lynchmaniac.poilight.entite;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Configure all the style for each side of the cell.
 * 
 * @author vpiard
 * @since 0.1
 */
public class BorderStyle {
	
	/**
	 * The bottom cell style
	 */
	private Short bottom = CellStyle.BORDER_NONE;
	/**
	 * The top cell style
	 */
	private Short top = CellStyle.BORDER_NONE;
	/**
	 * The left cell style
	 */
	private Short left = CellStyle.BORDER_NONE;
	/**
	 * The right cell style
	 */
	private Short right = CellStyle.BORDER_NONE;
	
	/**
	 * Constructor
	 */
	public BorderStyle() {
		super();
	}
	
	/**
	 * Constructor.
	 * 
	 * @param bottom the bottom cell style
	 * @param top the top cell style
	 * @param left the left cell style
	 * @param right the right cell style
	 */
	public BorderStyle(Short bottom, Short top, Short left, Short right) {
		super();
		this.bottom = bottom;
		this.top = top;
		this.left = left;
		this.right = right;
	}
	/**
	 * @return the bottom
	 */
	public Short getBottom() {
		return bottom;
	}
	/**
	 * @param bottom the bottom to set
	 */
	public void setBottom(Short bottom) {
		this.bottom = bottom;
	}
	/**
	 * @return the top
	 */
	public Short getTop() {
		return top;
	}
	/**
	 * @param top the top to set
	 */
	public void setTop(Short top) {
		this.top = top;
	}
	/**
	 * @return the left
	 */
	public Short getLeft() {
		return left;
	}
	/**
	 * @param left the left to set
	 */
	public void setLeft(Short left) {
		this.left = left;
	}
	/**
	 * @return the right
	 */
	public Short getRight() {
		return right;
	}
	/**
	 * @param right the right to set
	 */
	public void setRight(Short right) {
		this.right = right;
	}
	
	/**
	 * Save the configuration of the cell's border.
	 *  
	 * @param style the configuration's style
	 */
	public void getBorderStyle(TableStyle style) {
		style.setCellBorderBottom(getBottom());
		style.setCellBorderTop(getTop());
		style.setCellBorderLeft(getLeft());
		style.setCellBorderRight(getRight());
	}
}
