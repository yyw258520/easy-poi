package wsepr.easypoi.excel.style.font;

import org.apache.poi.hssf.usermodel.HSSFFont;

/**
 * 上标，下标
 * 
 */
public enum TypeOffset {
	/**
	 * 正常
	 */
	NONE(HSSFFont.SS_NONE), 
	/**
	 * 上标
	 */
	SUPER(HSSFFont.SS_SUPER),
	/**
	 * 下标
	 */
	SUB(HSSFFont.SS_SUB);

	private short offset;

	private TypeOffset(short offset) {
		this.offset = offset;
	}

	public short getOffset() {
		return offset;
	}

}
