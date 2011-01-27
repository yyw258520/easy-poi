package wsepr.easypoi.excel;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public final class ExcelContext {
	private Map<HSSFSheet, HSSFPatriarch> patriarchCache = new HashMap<HSSFSheet, HSSFPatriarch>();
	private HSSFWorkbook workBook;
	private HSSFCellStyle tempCellStyle;// 临时的样式
	private HSSFFont tempFont;// 临时的字体
	private HSSFSheet workingSheet;
	/**
	 * 默认样式
	 */
	private DefaultExcelStyle defaultStyle;
	/**
	 * 当前选择的工作簿
	 */
	private int workingSheetIndex = 0;
	
	protected ExcelContext(){};
	
	public HSSFWorkbook getWorkBook() {
		return workBook;
	}

	public void setWorkBook(HSSFWorkbook workBook) {
		this.workBook = workBook;
	}

	public HSSFCellStyle getTempCellStyle() {
		return tempCellStyle;
	}

	public void setTempCellStyle(HSSFCellStyle tempCellStyle) {
		this.tempCellStyle = tempCellStyle;
	}

	public HSSFFont getTempFont() {
		return tempFont;
	}

	public void setTempFont(HSSFFont tempFont) {
		this.tempFont = tempFont;
	}

	public HSSFSheet getWorkingSheet() {
		return workingSheet;
	}

	public void setWorkingSheet(HSSFSheet workingSheet) {
		this.workingSheet = workingSheet;
		this.workingSheetIndex = this.workBook.getSheetIndex(workingSheet);
	}

	/**
	 * 返回Patriarch，每个工作表对有一个Patriarch，Patriarch是所有图形的容器
	 * 
	 * @return
	 */
	public HSSFPatriarch getHSSFPatriarch(HSSFSheet sheet) {
		HSSFPatriarch patr = null;
		try {
			patr = patriarchCache.get(sheet);
			if (patr == null) {
				patr = sheet.createDrawingPatriarch();
				patriarchCache.put(sheet, patr);
			}
		} catch (Exception e) {
			patr = sheet.createDrawingPatriarch();
		}
		return patr;
	}

	public void setDefaultStyle(DefaultExcelStyle defaultStyle) {
		this.defaultStyle = defaultStyle;
	}

	public DefaultExcelStyle getDefaultStyle() {
		return defaultStyle;
	}

	public int getWorkingSheetIndex() {
		return workingSheetIndex;
	}
}
