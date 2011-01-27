package wsepr.easypoi.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * 默认样式
 * 
 * @author luxiaopan
 * 
 */
public class DefaultExcelStyle {
	private short fontSize = 12;

	private String fontName = "宋体";

	private short backgroundColor = HSSFColor.AUTOMATIC.index;

	private short fillPattern = HSSFCellStyle.NO_FILL;

	private short align = HSSFCellStyle.ALIGN_GENERAL;

	private short vAlign = HSSFCellStyle.VERTICAL_CENTER;

	private short fontColor = HSSFFont.COLOR_NORMAL;

	private String defaultDatePattern = "yyyy/MM/dd HH:mm:ss";
	
	/**
	 * 字体大小，默认12
	 * @param fontSize
	 */
	public void setFontSize(short fontSize) {
		this.fontSize = fontSize;
	}

	/**
	 * 字体大小，默认12
	 * @param fontSize
	 */
	public short getFontSize() {
		return fontSize;
	}

	/**
	 * 字体名称，默认“宋体”
	 * @param fontName
	 */
	public void setFontName(String fontName) {
		this.fontName = fontName;
	}

	/**
	 * 字体名称，默认“宋体”
	 */
	public String getFontName() {
		return fontName;
	}

	/**
	 * 背景色，默认HSSFColor.AUTOMATIC.index，即无填充色
	 * @param backgroundColor 颜色，例如HSSFColor.RED.index
	 */
	public void setBackgroundColor(short backgroundColor) {
		this.backgroundColor = backgroundColor;
	}

	/**
	 * 背景色，默认HSSFColor.AUTOMATIC.index，即无色
	 */
	public short getBackgroundColor() {
		return backgroundColor;
	}

	/**
	 * 背景色填充样式，默认HSSFCellStyle.NO_FILL，即无填充
	 * 要设置填充色必须把该值设置为NO_FILL以外的值
	 * @param fillPattern 
	 */
	public void setFillPattern(short fillPattern) {
		this.fillPattern = fillPattern;
	}

	/**
	 * 背景色填充样式，默认HSSFCellStyle.NO_FILL，即无填充
	 * 要设置填充色必须把该值设置为NO_FILL以外的值
	 */
	public short getFillPattern() {
		return fillPattern;
	}

	/**
	 * 水平对齐方式，默认是HSSFCellStyle.ALIGN_GENERAL
	 * @param align 对齐方式，例如HSSFCellStyle.ALIGN_CENTER。可选值：ALIGN_GENERAL, ALIGN_LEFT, ALIGN_CENTER, ALIGN_RIGHT, ALIGN_FILL, ALIGN_JUSTIFY, ALIGN_CENTER_SELECTION
	 */
	public void setAlign(short align) {
		this.align = align;
	}

	/**
	 * 水平对齐方式，默认是HSSFCellStyle.ALIGN_GENERAL
	 */
	public short getAlign() {
		return align;
	}

	/**
	 * 垂直对齐方式，默认HSSFCellStyle.VERTICAL_CENTER，即居中
	 * @param vAlign 对齐方式，例如HSSFCellStyle.VERTICAL_TOP。可选值：VERTICAL_TOP, VERTICAL_CENTER, VERTICAL_BOTTOM, VERTICAL_JUSTIFY
	 */
	public void setVAlign(short vAlign) {
		this.vAlign = vAlign;
	}

	/**
	 * 垂直对齐方式，默认HSSFCellStyle.VERTICAL_CENTER，即居中
	 * @param vAlign
	 */
	public short getVAlign() {
		return vAlign;
	}

	/**
	 * 字体颜色，默认是HSSFFont.COLOR_NORMAL
	 * @param fontColor 颜色，例如HSSFColor.RED.index
	 */
	public void setFontColor(short fontColor) {
		this.fontColor = fontColor;
	}

	/**
	 * 字体颜色，默认是HSSFFont.COLOR_NORMAL
	 */
	public short getFontColor() {
		return fontColor;
	}
	
	/**
	 * 设置默认的日期格式化模式，默认是yyyy/MM/dd HH:mm:ss
	 * @param defaultDatePattern
	 */
	public void setDefaultDatePattern(String defaultDatePattern) {
		this.defaultDatePattern = defaultDatePattern;
	}

	/**
	 * 返回默认的日期格式化模式，默认是yyyy/MM/dd HH:mm:ss
	 * @return
	 */
	public String getDefaultDatePattern() {
		return defaultDatePattern;
	}

}
