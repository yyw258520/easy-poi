package wsepr.easypoi.excel.style.font;

import org.apache.poi.hssf.usermodel.HSSFFont;

/**
 * 字体加粗
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

	/**
	 * 根据值返回对应的枚举值
	 * @param weight
	 * @return
	 */
	public static BoldWeight instance(short weight){
		for(BoldWeight e : BoldWeight.values()){
			if(e.getWeight() == weight){
				return e;
			}
		}
		return BoldWeight.NORMAL;
	}
}
