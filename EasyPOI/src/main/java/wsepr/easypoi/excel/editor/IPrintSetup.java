package wsepr.easypoi.excel.editor;

import org.apache.poi.hssf.usermodel.HSSFPrintSetup;

/**
 * 设置打印格式
 * @author luxiaopan
 *
 */
public interface IPrintSetup {
	
	public void setup(HSSFPrintSetup printSetup);
}
