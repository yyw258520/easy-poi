package wsepr.easypoi.excel;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 存放公共变量
 * @author q
 *
 */
public final class ExcelContext {
	/**
	 * 缓存style对象
	 */
	private Map<Integer, HSSFCellStyle> styleCache = new HashMap<Integer, HSSFCellStyle>();
	private Map<Integer, HSSFFont> fontCache = new HashMap<Integer, HSSFFont>();
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
	
	protected ExcelContext(HSSFWorkbook workBook){
		this.workBook = workBook;
		short numStyle = workBook.getNumCellStyles();
		for(short i=0; i<numStyle;i++){
			HSSFCellStyle style = workBook.getCellStyleAt(i);
			if(style != tempCellStyle){
				styleCache.put(style.hashCode() - style.getIndex(), style);
			}
		}
		short numFont = workBook.getNumberOfFonts();
		for(short i=0; i<numFont;i++){
			HSSFFont font = workBook.getFontAt(i);
			if(font != tempFont){
				fontCache.put(font.hashCode() - font.getIndex(), font);
			}
		}
	};
	
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
		workingSheetIndex = workBook.getSheetIndex(workingSheet);
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

	public void setStyleCache(Map<Integer, HSSFCellStyle> styleCache) {
		this.styleCache = styleCache;
	}

	public Map<Integer, HSSFCellStyle> getStyleCache() {
		return styleCache;
	}

	public void setFontCache(Map<Integer, HSSFFont> fontCache) {
		this.fontCache = fontCache;
	}

	public Map<Integer, HSSFFont> getFontCache() {
		return fontCache;
	}
}
