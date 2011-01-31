package wsepr.easypoi.excel.style.font;

import org.apache.poi.hssf.usermodel.HSSFFont;

import wsepr.easypoi.excel.style.Color;

public class Font {
	private HSSFFont font;

	public Font(HSSFFont font) {
		this.font = font;
	}

	public Font boldweight(BoldWeight boldweight) {
		font.setBoldweight(boldweight.getWeight());
		return this;
	}

	public Font charSet(CharSet charset) {
		font.setCharSet(charset.getCharset());
		return this;
	}

	public Font color(Color color) {
		if(color.equals(Color.AUTOMATIC)){
			font.setColor(HSSFFont.COLOR_NORMAL);
		}else{
			font.setColor(color.getIndex());
		}
		return this;
	}

	public Font fontHeight(int height) {
		font.setFontHeight((short)height);
		return this;
	}

	public Font fontHeightInPoints(int height) {
		font.setFontHeightInPoints((short)height);
		return this;
	}

	public Font fontName(String name) {
		font.setFontName(name);
		return this;
	}

	public Font italic(boolean italic) {
		font.setItalic(italic);
		return this;
	}

	public Font strikeout(boolean strikeout) {
		font.setStrikeout(strikeout);
		return this;
	}

	public Font typeOffset(TypeOffset offset) {
		font.setTypeOffset(offset.getOffset());
		return this;
	}

	public Font underline(Underline underline) {
		font.setUnderline(underline.getLine());
		return this;
	}
	
	//TODO 加上get方法
}
