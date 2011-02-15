package wsepr.easypoi.excel.style.font;

import org.apache.poi.hssf.usermodel.HSSFFont;

/**
 * 字符集
 *
 */
public enum CharSet {
	ANSI(HSSFFont.ANSI_CHARSET), 
	
	DEFAULT(HSSFFont.DEFAULT_CHARSET), 
	
	SYMBOL(HSSFFont.SYMBOL_CHARSET);

	private byte charset;

	private CharSet(byte charset){
		this.charset = charset;
	}

	public byte getCharset() {
		return charset;
	}
	
	/**
	 * 根据值返回对应的枚举值
	 * @param weight
	 * @return
	 */
	public static CharSet instance(byte charset){
		for(CharSet e : CharSet.values()){
			if(e.getCharset() == charset){
				return e;
			}
		}
		return CharSet.DEFAULT;
	}

}
