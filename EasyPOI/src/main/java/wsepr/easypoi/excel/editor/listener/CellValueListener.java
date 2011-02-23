package wsepr.easypoi.excel.editor.listener;

import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.editor.CellEditor;

/**
 *单元格内容设置监听器
 */
public interface CellValueListener {
	
	/**
	 * 实现该方法，在单元格内容改变时触发
	 * @param cell 目标单元格
	 * @param newValue 新值
	 * @param row 单元格所在的行
	 * @param col 单元格所在的列
	 * @param sheet 单元格所属的sheet
	 * @param excel Excel对象
	 */
	public void onValueChange(CellEditor cell, Object newValue, int row, int col, int sheetIndex, Excel excel);
}
