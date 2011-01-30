package wsepr.easypoi.excel.editor;

import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import wsepr.easypoi.excel.ExcelContext;
import wsepr.easypoi.excel.util.ExcelUtil;


public class ExcelSheetEditor extends AbstractExcelEditor{

	private HSSFSheet sheet;
	private int sheetIndex;
	public ExcelSheetEditor(HSSFSheet sheet, ExcelContext context){
		super(context);
		this.sheet = sheet;
		this.sheetIndex = this.workBook.getSheetIndex(this.sheet);
	}
	
	/**
	 * 设置页眉
	 * @param left
	 * @param center
	 * @param right
	 * @return
	 */
	public ExcelSheetEditor header(String left, String center, String right){
		HSSFHeader header = this.sheet.getHeader();
		header.setLeft(left == null ? "" : left);
		header.setCenter(center == null ? "" : center);
		header.setRight(right == null ? "" : right);
		return this;
	}
	
	/**
	 * 设置页脚
	 * @param left
	 * @param center
	 * @param right
	 * @return
	 */
	public ExcelSheetEditor footer(String left, String center, String right){
		HSSFFooter footer = this.sheet.getFooter();
		footer.setLeft(left == null ? "" : left);
		footer.setCenter(center == null ? "" : center);
		footer.setRight(right == null ? "" : right);
		return this;
	}
	
	/**
	 * 设置工作表名
	 * @param name 表名
	 * @return
	 */
	public ExcelSheetEditor sheetName(String name){
		this.workBook.setSheetName(this.sheetIndex, name);
		return this;
	}
	
	/**
	 * 把该表设置为活动状态，用Excel打开后首先显示该表
	 * @return
	 */
	public ExcelSheetEditor active(){
		this.workBook.setActiveSheet(this.sheetIndex);
		return this;
	}
	
	/**
	 * 冻结行和列
	 * @param row 从指定的行开始冻结，如果传入0则不冻结任何行
	 * @param col 从指定的列开始冻结，如果传入0则不冻结任何列
	 * @return
	 */
	public ExcelSheetEditor freeze(int row, int col){
		if(row < 0){
			row = 0;
		}
		if(col < 0){
			col = 0;
		}
		this.sheet.createFreezePane(col, row, col, row);
		return this;
	}
	
	/**
	 * 获取工作表的行数
	 * @return 表行数
	 */
	public int getLastRowNum(){
		return ExcelUtil.getLastRowNum(sheet);
	}
	
	/**
	 * 是否显示表格线
	 * @param show
	 * @return
	 */
	public ExcelSheetEditor displayGridlines(boolean show){
		this.sheet.setDisplayGridlines(show);
		return this;
	}
	
	/**
	 * 是否打印表格线
	 * @param newPrintGridlines
	 * @return
	 */
	public ExcelSheetEditor printGridlines(boolean newPrintGridlines){
		this.sheet.setPrintGridlines(newPrintGridlines);
		return this;
	}
	
	/**
	 * 设置是否适合页面大小
	 * @param isFit
	 * @return
	 */
	public ExcelSheetEditor fitToPage(boolean isFit){
		this.sheet.setFitToPage(isFit);
		return this;
	}
	
	/**
	 * 设置打印时内容是否水平居中
	 * @param isCenter
	 * @return
	 */
	public ExcelSheetEditor horizontallyCenter(boolean isCenter){
		this.sheet.setHorizontallyCenter(isCenter);
		return this;
	}
}
