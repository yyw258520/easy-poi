package wsepr.easypoi.excel.editor;

import java.util.List;

import org.apache.poi.hssf.util.CellRangeAddress;

import wsepr.easypoi.excel.ExcelContext;


public abstract class AbstractRegionEditor<T> extends AbstractExcelEditor{

	protected AbstractRegionEditor(ExcelContext context) {
		super(context);
	}

	/**
	 * 新建一个单元格编辑器，包含所有单元格
	 * @return
	 */
	abstract protected ExcelCellEditor newCellEditor();
	
	/**
	 * 新建一个单元格编辑器，包含上边单元格
	 * @return
	 */
	abstract protected ExcelCellEditor newTopCellEditor();
	
	/**
	 * 新建一个单元格编辑器，包含下边单元格
	 * @return
	 */
	abstract protected ExcelCellEditor newBottomCellEditor();
	
	/**
	 * 新建一个单元格编辑器，包含左边单元格
	 * @return
	 */
	abstract protected ExcelCellEditor newLeftCellEditor();
	
	/**
	 * 新建一个单元格编辑器，包含右边单元格
	 * @return
	 */
	abstract protected ExcelCellEditor newRightCellEditor();
	
	abstract protected List<CellRangeAddress> getCellRange();
	
	/**
	 * 设置外部四条边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public T borderOuter(int borderStyle, int borderColor){
		this.borderBottom(borderStyle, borderColor);
		this.borderLeft(borderStyle, borderColor);
		this.borderRight(borderStyle, borderColor);
		this.borderTop(borderStyle, borderColor);
		return (T) this;
	}
	
	/**
	 * 设置全部内外边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public T borderFull(int borderStyle, int borderColor){
		ExcelCellEditor cellEditor = newCellEditor();
		cellEditor.border(borderStyle, borderColor);
		return (T) this;
	}
	
	/**
	 * 设置外部左边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public T borderLeft(int borderStyle, int borderColor){
		//左边框
		ExcelCellEditor cellEditorLeft = this.newLeftCellEditor();
		cellEditorLeft.borderLeft(borderStyle, borderColor);
		return (T) this;
	}

	/**
	 * 设置外部右边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public T borderRight(int borderStyle, int borderColor){
		//右边框
		ExcelCellEditor cellEditorRight = this.newRightCellEditor();
		cellEditorRight.borderRight(borderStyle, borderColor);
		return (T) this;
	}
	
	/**
	 * 设置外部上边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public T borderTop(int borderStyle, int borderColor){
		//上边框
		ExcelCellEditor cellEditorTop = this.newTopCellEditor();
		cellEditorTop.borderTop(borderStyle, borderColor);
		return (T) this;
	}
	
	/**
	 * 设置外部下边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public T borderBottom(int borderStyle, int borderColor){
		//下边框
		ExcelCellEditor cellEditorBottom = this.newBottomCellEditor();
		cellEditorBottom.borderBottom(borderStyle, borderColor);
		return (T) this;
	}
	
	/**
	 * 设置字体
	 * 
	 * @param fontEditor
	 *            实现IFontEditor接口
	 * @return
	 */
	public T font(IFontEditor fontEditor) {
		ExcelCellEditor cellEditor = newCellEditor();
		cellEditor.font(fontEditor);
		return (T) this;
	}

	/**
	 * 设置背景色
	 * 
	 * @param bg
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public T bgColor(short bg) {
		ExcelCellEditor cellEditor = newCellEditor();
		cellEditor.bgColor(bg);
		return (T) this;
	}
	
	/**
	 * 设置背景色，可设置填充的样式
	 * 
	 * @param bg
	 *            颜色，例如HSSFColor.RED.index
	 * @param fillPattern
	 *            填充样式，如HSSFCellStyle.FINE_DOTS。可选值：NO_FILL,
	 *            SOLID_FOREGROUND, FINE_DOTS, ALT_BARS, SPARSE_DOTS,
	 *            THICK_HORZ_BANDS, THICK_VERT_BANDS, THICK_BACKWARD_DIAG,
	 *            THICK_FORWARD_DIAG, BIG_SPOTS, BRICKS, THIN_HORZ_BANDS,
	 *            THIN_VERT_BANDS, THIN_BACKWARD_DIAG, THIN_FORWARD_DIAG,
	 *            SQUARES, DIAMONDS
	 * @return
	 */
	public T bgColor(short bg, short fillPattern) {
		ExcelCellEditor cellEditor = newCellEditor();
		cellEditor.bgColor(bg, fillPattern);
		return (T) this;
	}
	
	/**
	 * 设置水平对齐方式
	 * @param align 对齐方式，例如HSSFCellStyle.ALIGN_CENTER。可选值：ALIGN_GENERAL, ALIGN_LEFT, ALIGN_CENTER, ALIGN_RIGHT, ALIGN_FILL, ALIGN_JUSTIFY, ALIGN_CENTER_SELECTION
	 * @return
	 */
	public T align(Short align){
		ExcelCellEditor cellEditor = newCellEditor();
		cellEditor.align(align);
		return (T) this;
	}
	
	/**
	 * 设置垂直对齐方式
	 * @param align 对齐方式，例如HSSFCellStyle.VERTICAL_CENTER。可选值：VERTICAL_TOP, VERTICAL_CENTER, VERTICAL_BOTTOM, VERTICAL_JUSTIFY
	 * @return
	 */
	public T vAlign(short align){
		ExcelCellEditor cellEditor = newCellEditor();
		cellEditor.vAlign(align);
		return (T) this;
	}
	
	/**
	 * 是否自动换行。
	 * @param autoWarp true自动换行，false不换行
	 * @return
	 */
	public T warpText(boolean autoWarp){
		ExcelCellEditor cellEditor = newCellEditor();
		cellEditor.warpText(autoWarp);
		return (T) this;
	}
	
	/**
	 * 合并区间，注意：合并区间可能导致区间内一些单元格的值丢失
	 * 
	 * @return
	 */
	public T merge() {
		for(CellRangeAddress cellRange : this.getCellRange()){
			this.workingSheet.addMergedRegion(cellRange);
		}
		return (T) this;
	}
}
