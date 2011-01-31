package wsepr.easypoi.excel.style.font;

import org.apache.poi.hssf.usermodel.HSSFFont;

/**
 * 字体下横线样式
 *
 */
public enum Underline {
	NONE(HSSFFont.U_NONE), 
	
	/**
	 * 单下横线
	 */
	SINGLE(HSSFFont.U_SINGLE), 
	
	/**
	 * 双下横线
	 */
	DOUBLE(HSSFFont.U_DOUBLE), 
	
	/**
	 * 会计用单下横线
	 */
	SINGLE_ACCOUNTING(HSSFFont.U_SINGLE_ACCOUNTING), 
	
	/**
	 * 会计用双下横线
	 */
	DOUBLE_ACCOUNTING(HSSFFont.U_DOUBLE_ACCOUNTING);
	
	private byte line;

	private Underline(byte line){
		this.line = line;
	}

	public byte getLine() {
		return line;
	}
	
	


}
