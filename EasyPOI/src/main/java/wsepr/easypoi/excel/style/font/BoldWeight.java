package wsepr.easypoi.excel.style.font;

import org.apache.poi.hssf.usermodel.HSSFFont;

/**
 * 加粗
 */
public enum BoldWeight {
	/**
	 * 正常
	 */
	NORMAL(HSSFFont.BOLDWEIGHT_NORMAL), 
	
	/**
	 * 加粗
	 */
	BOLD(HSSFFont.BOLDWEIGHT_BOLD);

	private short weight;

	private BoldWeight(short weight) {
		this.weight = weight;
	}

	public short getWeight() {
		return weight;
	}		

}
