package wsepr.easypoi.excel.editor;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.util.CellRangeAddress;

import wsepr.easypoi.excel.ExcelContext;
import wsepr.easypoi.excel.util.ExcelUtil;


public class ColumnEditor extends AbstractRegionEditor<ColumnEditor> {

	private int col = 0;
	public ColumnEditor(int col, ExcelContext context) {
		super(context);
		this.col = col;
	}
	
	/**
	 * 设置该行的内容，该方法会覆盖该行已有的内容
	 * 
	 * @param rowData
	 *            内容数组，如果里面有Date类型的元素，则会用默认模式yyyy/MM/dd HH:mm:ss格式化。
	 *            使用Excel.setDefaultDatePattern方法设置默认模式
	 * @return
	 */
	public ColumnEditor value(Object[] rowData) {
		value(rowData, 0);
		return this;
	}

	/**
	 * 设置该行的内容，该方法会覆盖该行已有的内容
	 * 
	 * @param rowData
	 *            内容数组，如果里面有Date类型的元素，则会用默认模式yyyy/MM/dd HH:mm:ss格式化。
	 *            使用Excel.setDefaultDatePattern方法设置默认模式
	 * @param startCol
	 *            从指定的列开始写入，从0开始
	 * @return
	 */
	public ColumnEditor value(Object[] rowData, int startRow) {
		if (startRow < 0) {
			startRow = 0;
		}
		insertData(rowData, col, startRow);
		return this;
	}
	
	/**
	 * 设置列的宽度
	 * @param width 要设置的宽度。1表示一个文字宽度的1/256
	 */
	public ColumnEditor width(int width){
		workingSheet.setColumnWidth(col, width);
		return this;
	}
	
	/**
	 * 在原来宽度基础上增加列的宽度
	 * @param width 要设置的宽度。1表示一个文字宽度的1/256
	 */
	public ColumnEditor addWidth(int width){
		int w = workingSheet.getColumnWidth(col);
		workingSheet.setColumnWidth(col, width+w);
		return this;
	}
	
	/**
	 * 根据内容自动设置列宽度。自动计算宽度性能比较低，因此建议在操作完数据后调用一次
	 */
	public void autoWidth(){
		workingSheet.autoSizeColumn((short)col, false);
		workingSheet.setColumnWidth(col, workingSheet.getColumnWidth(col)+1000);
	}
	
	/**
	 * 获取该列的第row行的单元格
	 * @param row 列，从0开始
	 * @return
	 */
	public CellEditor cell(int row){
		CellEditor cellEditor = new CellEditor(row, col, ctx);
		return cellEditor;
	}
	
	/**
	 * 插入数据
	 * 
	 * @param rowData
	 *            待插入的数据
	 * @param col
	 *            列序号，从0开始
	 * @param startRow
	 *            开始插入的列，从0开始
	 * @throws Exception
	 */
	private void insertData(Object[] rowData, int col, int startRow) {
		short i = 0;
		for (Object obj : rowData) {
			CellEditor cellEditor = new CellEditor(startRow + i, col, ctx);
			cellEditor.value(obj);
			i++;
		}
	}
	
	@Override
	protected CellEditor newBottomCellEditor() {
		int lastRowNum = ExcelUtil.getLastRowNum(workingSheet);
		CellEditor cellEditor = new CellEditor(ctx);
		cellEditor.add(lastRowNum, col);
		return cellEditor;
	}

	@Override
	protected CellEditor newCellEditor() {
		CellEditor cellEditor = new CellEditor(ctx);
		int lastRowNum = ExcelUtil.getLastRowNum(workingSheet);
		int firstRowNum = workingSheet.getFirstRowNum();
		for(int i=firstRowNum; i<= lastRowNum; i++){
			HSSFRow row = getRow(i);
			cellEditor.add(row.getRowNum(), col);
		}
		return cellEditor;
	}

	@Override
	protected CellEditor newLeftCellEditor() {
		return newCellEditor();
	}

	@Override
	protected CellEditor newRightCellEditor() {
		return newCellEditor();
	}

	@Override
	protected CellEditor newTopCellEditor() {
		int firstRowNum = workingSheet.getFirstRowNum();
		CellEditor cellEditor = new CellEditor(ctx);
		cellEditor.add(firstRowNum, col);
		return cellEditor;
	}

	@Override
	protected List<CellRangeAddress> getCellRange() {
		int firstRowNum = workingSheet.getFirstRowNum();
		int lastRowNum = ExcelUtil.getLastRowNum(workingSheet);
		List<CellRangeAddress> cellRangeList = new ArrayList<CellRangeAddress>();
		cellRangeList.add(new CellRangeAddress(firstRowNum, lastRowNum, col, col));
		return cellRangeList;
	}

	protected int getCol() {
		return col;
	}	
}