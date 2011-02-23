package wsepr.easypoi.excel.editor;

import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import wsepr.easypoi.excel.ExcelContext;
import wsepr.easypoi.excel.editor.listener.CellValueListener;
import wsepr.easypoi.excel.util.ExcelUtil;

/**
 * 表单编辑器
 *
 */
public class SheetEditor extends AbstractEditor{

	private HSSFSheet sheet;
	private int sheetIndex;
	public SheetEditor(HSSFSheet sheet, ExcelContext context){
		super(context);
		this.sheet = sheet;
		sheetIndex = workBook.getSheetIndex(this.sheet);
	}
	
	/**
	 * 设置页眉
	 * @param left
	 * @param center
	 * @param right
	 * @return
	 */
	public SheetEditor header(String left, String center, String right){
		HSSFHeader header = sheet.getHeader();
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
	public SheetEditor footer(String left, String center, String right){
		HSSFFooter footer = sheet.getFooter();
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
	public SheetEditor sheetName(String name){
		workBook.setSheetName(sheetIndex, name);
		return this;
	}
	
	/**
	 * 把该表设置为活动状态，用Excel打开后首先显示该表
	 * @return
	 */
	public SheetEditor active(){
		workBook.setActiveSheet(sheetIndex);
		return this;
	}
	
	/**
	 * 冻结行和列
	 * @param row 从指定的行开始冻结，如果传入0则不冻结任何行
	 * @param col 从指定的列开始冻结，如果传入0则不冻结任何列
	 * @return
	 */
	public SheetEditor freeze(int row, int col){
		if(row < 0){
			row = 0;
		}
		if(col < 0){
			col = 0;
		}
		sheet.createFreezePane(col, row, col, row);
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
	public SheetEditor displayGridlines(boolean show){
		sheet.setDisplayGridlines(show);
		return this;
	}
	
	/**
	 * 是否打印表格线
	 * @param newPrintGridlines
	 * @return
	 */
	public SheetEditor printGridlines(boolean newPrintGridlines){
		sheet.setPrintGridlines(newPrintGridlines);
		return this;
	}
	
	/**
	 * 设置是否适合页面大小
	 * @param isFit
	 * @return
	 */
	public SheetEditor fitToPage(boolean isFit){
		sheet.setFitToPage(isFit);
		return this;
	}
	
	/**
	 * 设置打印时内容是否水平居中
	 * @param isCenter
	 * @return
	 */
	public SheetEditor horizontallyCenter(boolean isCenter){
		sheet.setHorizontallyCenter(isCenter);
		return this;
	}
	
	/**
	 * 保护工作表
	 * @param pw 密码
	 * @return
	 */
	public SheetEditor password(String pw){
		sheet.protectSheet(pw);
		return this;
	}
	
	/**
	 * 详细设置打印属性
	 * @param printSetup
	 * @return
	 */
	public SheetEditor printSetup(IPrintSetup printSetup){
		printSetup.setup(sheet.getPrintSetup());
		return this;
	}
	
	/**
	 * 只有设置为true，printSetup中的setFitHeight和setFitWidth才会生效
	 * @param b
	 * @return
	 */
	public SheetEditor autobreaks(boolean b){
		sheet.setAutobreaks(b);
		return this;
	}
	
	/**
	 * 添加单元格监听器
	 * @param listener 监听器，在单元格的值改变时触发
	 */
	public void addCellValueListener(CellValueListener listener){
		ctx.getListenerList(sheetIndex).add(listener);
	}
	
	/**
	 * 移除单元格监听器
	 * @param listener 监听器，在单元格的值改变时触发
	 */
	public void removeCellValueListener(CellValueListener listener){
		ctx.getListenerList(sheetIndex).remove(listener);
	}
}
