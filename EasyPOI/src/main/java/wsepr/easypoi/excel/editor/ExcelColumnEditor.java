package wsepr.easypoi.excel.editor;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.util.CellRangeAddress;

import wsepr.easypoi.excel.ExcelContext;
import wsepr.easypoi.excel.util.ExcelUtil;


public class ExcelColumnEditor extends AbstractRegionEditor<ExcelColumnEditor> {

	private List<Integer> workingCol = new ArrayList<Integer>();
	
	public ExcelColumnEditor(int col, ExcelContext context) {
		super(context);
		this.add(col);
	}

	/**
	 * 把更多的列加入编辑队列，以进行一系列相同的操作。该方法应该在改变列属性前调用
	 *	否则所做的修改不会影响到新加入的列
	 * @param col 列序号，从0开始
	 * @param cols n个列序号，从0开始
	 * @return
	 */
	public ExcelColumnEditor add(int col, int... cols){
		if(col < 0){
			col = 0;
		}
		this.workingCol.add(col);
		for(int c : cols){
			if(c < 0){
				c = 0;
			}
			this.workingCol.add(c);
		}
		return this;
	}
	
	/**
	 * 设置该行的内容，该方法会覆盖该行已有的内容
	 * 
	 * @param rowData
	 *            内容数组，如果里面有Date类型的元素，则会用默认模式yyyy/MM/dd HH:mm:ss格式化。
	 *            使用Excel.setDefaultDatePattern方法设置默认模式
	 * @return
	 */
	public ExcelColumnEditor value(Object[] rowData) {
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
	public ExcelColumnEditor value(Object[] rowData, int startRow) {
		if (startRow < 0) {
			startRow = 0;
		}
		for (Integer col : workingCol) {
			insertData(rowData, col, startRow);
		}
		return this;
	}
	
	/**
	 * 设置列的宽度
	 * @param width 要设置的宽度。1表示一个文字宽度的1/256
	 */
	public ExcelColumnEditor width(int width){
		for (Integer col : workingCol) {
			this.workingSheet.setColumnWidth(col, width);
		}
		return this;
	}
	
	/**
	 * 在原来宽度基础上增加列的宽度
	 * @param width 要设置的宽度。1表示一个文字宽度的1/256
	 */
	public ExcelColumnEditor addWidth(int width){
		int w = 0;
		for (Integer col : workingCol) {
			w = this.workingSheet.getColumnWidth(col);
			this.workingSheet.setColumnWidth(col, width+w);
		}
		return this;
	}
	
	/**
	 * 根据内容自动设置列宽度。自动计算宽度性能比较低，因此建议在操作完数据后调用一次
	 */
	public void autoWidth(){
		for (Integer col : workingCol) {
			this.workingSheet.autoSizeColumn(col.shortValue(), false);
			this.workingSheet.setColumnWidth(col, this.workingSheet.getColumnWidth(col)+1000);
		}
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
			ExcelCellEditor cellEditor = new ExcelCellEditor(startRow + i, col, this.ctx);
			cellEditor.value(obj);
			i++;
		}
	}
	
	@Override
	protected ExcelCellEditor newBottomCellEditor() {
		int lastRowNum = ExcelUtil.getLastRowNum(this.workingSheet);
		ExcelCellEditor cellEditor = new ExcelCellEditor(this.ctx);
		for (Integer col : workingCol) {
			cellEditor.add(lastRowNum, col);
		}
		return cellEditor;
	}

	@Override
	protected ExcelCellEditor newCellEditor() {
		ExcelCellEditor cellEditor = new ExcelCellEditor(this.ctx);
//		for (Iterator<HSSFRow> rit = this.workingSheet.rowIterator(); rit.hasNext(); ) {
//			HSSFRow row = rit.next();
//			for (Integer col : workingCol) {
//				cellEditor.add(row.getRowNum(), col);
//			}
//		}
		int lastRowNum = ExcelUtil.getLastRowNum(this.workingSheet);
		int firstRowNum = this.workingSheet.getFirstRowNum();
		for(int i=firstRowNum; i<= lastRowNum; i++){
			HSSFRow row = getRow(i);
			for (Integer col : workingCol) {
				cellEditor.add(row.getRowNum(), col);
			}
		}
		return cellEditor;
	}

	@Override
	protected ExcelCellEditor newLeftCellEditor() {
		return newCellEditor();
	}

	@Override
	protected ExcelCellEditor newRightCellEditor() {
		return newCellEditor();
	}

	@Override
	protected ExcelCellEditor newTopCellEditor() {
		int firstRowNum = this.workingSheet.getFirstRowNum();
		ExcelCellEditor cellEditor = new ExcelCellEditor(this.ctx);
		for (Integer col : workingCol) {
			cellEditor.add(firstRowNum, col);
		}
		return cellEditor;
	}

	@Override
	protected List<CellRangeAddress> getCellRange() {
		int firstRowNum = this.workingSheet.getFirstRowNum();
		int lastRowNum = ExcelUtil.getLastRowNum(this.workingSheet);
		List<CellRangeAddress> cellRangeList = new ArrayList<CellRangeAddress>();
		for (Integer col : workingCol) {
			cellRangeList.add(new CellRangeAddress(firstRowNum, lastRowNum, col, col));
		}
		return cellRangeList;
	}	
}
