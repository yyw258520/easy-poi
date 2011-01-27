package wsepr.easypoi.excel.editor;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.util.CellRangeAddress;

import wsepr.easypoi.excel.ExcelContext;


public class ExcelRowEditor extends AbstractRegionEditor<ExcelRowEditor> {
	private List<HSSFRow> workingRow = new ArrayList<HSSFRow>(2);
	public ExcelRowEditor(int row, ExcelContext context) {
		super(context);
		this.add(row);
	}

	/**
	 * 把更多的行加入编辑队列，以进行一系列相同的操作。该方法应该在改变行属性前调用
	 *	否则所做的修改不会影响到新加入的行
	 * @param row
	 * @param rows
	 * @return
	 */
	public ExcelRowEditor add(int row, int... rows) {
		HSSFRow newRow = this.getRow(row);
		this.workingRow.add(newRow);
		for(int r : rows){
			newRow = this.getRow(r);
			this.workingRow.add(newRow);
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
	public ExcelRowEditor value(Object[] rowData) {
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
	public ExcelRowEditor value(Object[] rowData, int startCol) {
		if (startCol < 0) {
			startCol = 0;
		}
		for (HSSFRow row : workingRow) {
			insertData(rowData, row, startCol, true);
		}
		return this;
	}
	
	/**
	 * 插入一行，原来的内容会自动下移一行
	 * @param rowData 内容数组，如果里面有Date类型的元素，则会用默认模式yyyy/MM/dd HH:mm:ss格式化。
	 *            使用Excel.setDefaultDatePattern方法设置默认模式
	 * @return
	 */
	public ExcelRowEditor insert(Object[] rowData){
		return insert(rowData, 0);
	}
	
	/**
	 * 插入一行，原来的内容会自动下移一行
	 * @param rowData 内容数组，如果里面有Date类型的元素，则会用默认模式yyyy/MM/dd HH:mm:ss格式化。
	 *            使用Excel.setDefaultDatePattern方法设置默认模式
	 * @param startCol 从指定的列开始写入，从0开始
	 * @return
	 */
	public ExcelRowEditor insert(Object[] rowData, int startCol){
		if (startCol < 0) {
			startCol = 0;
		}
		for (HSSFRow row : workingRow) {
			insertData(rowData, row, startCol, false);
		}
		return this;
	}
	
	/**
	 * 在行末添加内容
	 * @param rowData
	 * @return
	 */
	public ExcelRowEditor append(Object[] rowData){
		for (HSSFRow row : workingRow) {
			insertData(rowData, row, row.getLastCellNum(), true);
		}
		return this;
	}
	
	
	/**
	 * 插入数据
	 * 
	 * @param rowData
	 *            待插入的数据
	 * @param row
	 *            行对象
	 * @param startCol
	 *            开始插入的列，从0开始
	 * @param overwrite
	 *            是否覆盖该行数据
	 * @throws Exception
	 */
	private void insertData(Object[] rowData, HSSFRow row, int startCol, boolean overwrite) {
		if (!overwrite) {
			this.workingSheet.shiftRows(row.getRowNum(), this.workingSheet.getLastRowNum(), 1, true, false);
		}
		short i = 0;
		for (Object obj : rowData) {
			ExcelCellEditor cellEditor = new ExcelCellEditor(row.getRowNum(), startCol + i, this.ctx);
			cellEditor.value(obj);
			i++;
		}
	}

	@Override
	protected ExcelCellEditor newCellEditor(){
		ExcelCellEditor cellEditor = new ExcelCellEditor(this.ctx);
		short minColIx = 0;
		short maxColIx = 0;
		for (HSSFRow row : workingRow) {
			minColIx = row.getFirstCellNum();
			maxColIx = row.getLastCellNum();
			for(int i=minColIx; i< maxColIx; i++){
				cellEditor.add(row.getRowNum(), i);
			}
		}
		return cellEditor;
	}
	
	@Override
	protected ExcelCellEditor newBottomCellEditor() {
		return newCellEditor();
	}

	@Override
	protected ExcelCellEditor newLeftCellEditor() {
		ExcelCellEditor cellEditor = new ExcelCellEditor(this.ctx);
		for (HSSFRow row : workingRow) {
			cellEditor.add(row.getRowNum(), row.getFirstCellNum());
		}
		return cellEditor;
	}

	@Override
	protected ExcelCellEditor newRightCellEditor() {
		ExcelCellEditor cellEditor = new ExcelCellEditor(this.ctx);
		for (HSSFRow row : workingRow) {
			cellEditor.add(row.getRowNum(), row.getLastCellNum());
		}
		return cellEditor;
	}

	@Override
	protected ExcelCellEditor newTopCellEditor() {
		return newCellEditor();
	}

	@Override
	protected List<CellRangeAddress> getCellRange() {
		List<CellRangeAddress> cellRangeList = new ArrayList<CellRangeAddress>();
		for (HSSFRow row : workingRow) {
			cellRangeList.add(new CellRangeAddress(row.getRowNum(), row.getRowNum(), row.getFirstCellNum(), row.getLastCellNum()));
		}
		return cellRangeList;
	}
}
