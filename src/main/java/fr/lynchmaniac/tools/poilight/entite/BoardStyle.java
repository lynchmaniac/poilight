/**
 * 
 */
package fr.lynchmaniac.tools.poilight.entite;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import fr.lynchmaniac.tools.poilight.PoiLightConst;
import fr.lynchmaniac.tools.poilight.helpers.StyleHelper;

/**
 * This object is used to store the many configuration of the Excel's styles.
 * This object properies is after apply to the real properties of POI.
 * 
 * @author vpiard
 * @since 0.1
 */
public class BoardStyle {
	
	/**
	 * The font color.
	 */
	private XSSFColor fontColor = StyleHelper.getColor(0, 0, 0);
	/**
	 * The color of the border of the cell.
	 */
	private XSSFColor borderColor = StyleHelper.getColor(0, 0, 0);
	/**
	 * The style of the background of the cell.
	 */
	private XSSFColor fillColor = null;
	/**
	 * The font name.
	 */
	private String fontName = PoiLightConst.EXCEL_2007_DEFAULT_FONT_NAME;
	/**
	 * Indicate if the font is Bold.
	 */
	private boolean isBold = false;
	/**
	 * The font size.
	 */
	private Short fontSize = new Short("11");
	/**
	 * Indicate the cell's alignment.
	 */
	private Short alignment = CellStyle.ALIGN_LEFT;
	
	/**
	 * The style of the border left of the cell.
	 */
	private Short cellBorderLeft = CellStyle.BORDER_MEDIUM;
	
	/**
	 * The style of the border right of the cell.
	 */
	private Short cellBorderRight = CellStyle.BORDER_MEDIUM;
	
	/**
	 * The style of the border top of the cell.
	 */
	private Short cellBorderTop = CellStyle.BORDER_THIN;

	/**
	 * The style of the border bottom of the cell.
	 */
	private Short cellBorderBottom = CellStyle.BORDER_THIN;

	
	/**
	 * Constructor.
	 */
	public BoardStyle() {
		super();
	}

	/**
	 * Constructor.
	 * 
	 * @param fillColor the background color
	 */
	public BoardStyle(XSSFColor fillColor) {
		super();
		this.fillColor = fillColor;
	}
	

	/**
	 * Constructor.
	 * 
	 * @param fontColor the font color
	 * @param fillColor the background color
	 * @param fontName the font name
	 * @param isBold indicate if the font is bold
	 * @param fontSize the font size
	 * @param alignment the cell's alignment
	 */
	public BoardStyle(XSSFColor fontColor, XSSFColor fillColor, String fontName,
			boolean isBold, Short fontSize, Short alignment) {
		super();
		this.fontColor = fontColor;
		this.fillColor = fillColor;
		this.fontName = fontName;
		this.isBold = isBold;
		this.fontSize = fontSize;
		this.alignment = alignment;
	}


	/**
	 * @return fontColor
	 */
	public XSSFColor getFontColor() {
		return fontColor;
	}


	/**
	 * @param fontColor colorFont to define
	 */
	public void setFontColor(XSSFColor fontColor) {
		this.fontColor = fontColor;
	}


	/**
	 * @return fillColor
	 */
	public XSSFColor getFillColor() {
		return fillColor;
	}


	/**
	 * @return the borderColor
	 */
	public XSSFColor getBorderColor() {
		return borderColor;
	}

	/**
	 * @param borderColor the borderColor to set
	 */
	public void setBorderColor(XSSFColor borderColor) {
		this.borderColor = borderColor;
	}

	/**
	 * @param fillColor the background color of the cell
	 */
	public void setFillColor(XSSFColor fillColor) {
		this.fillColor = fillColor;
	}


	/**
	 * @return fontName
	 */
	public String getFontName() {
		return fontName;
	}


	/**
	 * @param fontName fontName to define
	 */
	public void setFontName(String fontName) {
		this.fontName = fontName;
	}


	/**
	 * @return isBold
	 */
	public boolean isBold() {
		return isBold;
	}


	/**
	 * @param isBold isBold to define
	 */
	public void setBold(boolean isBold) {
		this.isBold = isBold;
	}


	/**
	 * @return fontSize
	 */
	public Short getFontSize() {
		return fontSize;
	}


	/**
	 * @param fontSize fontHeight to define
	 */
	public void setFontSize(Short fontSize) {
		this.fontSize = fontSize;
	}


	/**
	 * @return alignment
	 */
	public Short getAlignment() {
		return alignment;
	}


	/**
	 * @param alignment alignment to define
	 */
	public void setAlignment(Short alignment) {
		this.alignment = alignment;
	}

	/**
	 * @return cellBorderLeft
	 */
	public Short getCellBorderLeft() {
		return cellBorderLeft;
	}

	/**
	 * @param cellBorderLeft cellBorderLeft to define
	 */
	public void setCellBorderLeft(Short cellBorderLeft) {
		this.cellBorderLeft = cellBorderLeft;
	}

	/**
	 * @return cellBorderRight
	 */
	public Short getCellBorderRight() {
		return cellBorderRight;
	}

	/**
	 * @param cellBorderRight cellBorderRight to define
	 */
	public void setCellBorderRight(Short cellBorderRight) {
		this.cellBorderRight = cellBorderRight;
	}

	/**
	 * @return cellBorderTop
	 */
	public Short getCellBorderTop() {
		return cellBorderTop;
	}

	/**
	 * @param cellBorderTop cellBorderTop to define
	 */
	public void setCellBorderTop(Short cellBorderTop) {
		this.cellBorderTop = cellBorderTop;
	}

	/**
	 * @return cellBorderBottom
	 */
	public Short getCellBorderBottom() {
		return cellBorderBottom;
	}

	/**
	 * @param cellBorderBottom cellBorderBottom to define
	 */
	public void setCellBorderBottom(Short cellBorderBottom) {
		this.cellBorderBottom = cellBorderBottom;
	}

}
