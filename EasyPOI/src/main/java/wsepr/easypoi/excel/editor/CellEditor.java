package wsepr.easypoi.excel.editor;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;

import wsepr.easypoi.excel.ExcelContext;
import wsepr.easypoi.excel.editor.listener.CellValueListener;
import wsepr.easypoi.excel.style.Align;
import wsepr.easypoi.excel.style.BorderStyle;
import wsepr.easypoi.excel.style.Color;
import wsepr.easypoi.excel.style.FillPattern;
import wsepr.easypoi.excel.style.VAlign;
import wsepr.easypoi.excel.style.font.Font;


public class CellEditor extends AbstractEditor{
	private List<HSSFCell> workingCell = new ArrayList<HSSFCell>(2);

	public CellEditor(int row, int col, ExcelContext context) {
		this(context);
		this.add(row, col);
	}
	
	public CellEditor(ExcelContext context) {
		super(context);
	}

	
	/**
	 * 写入内容
	 * 
	 * @param value
	 * @return
	 */
	public CellEditor value(Object value) {
		for (HSSFCell cell : workingCell) {
			this.setCellValue(cell, value, null);
		}
		return this;
	}

	/**
	 * 写入日期
	 * 
	 * @param value
	 *            日期
	 * @param pattern
	 *            格式化字符串
	 * @return
	 */
	public CellEditor value(Date value, String pattern) {
		for (HSSFCell cell : workingCell) {
			this.setCellValue(cell, value, pattern);
		}
		return this;
	}

	/**
	 * 添加其他单元格到编辑队列，该方法应该在改变单元格属性前调用
	 *	否则所做的修改不会影响到新加入的单元格
	 * 
	 * @param row 第n行，从0开始
	 * @param col 第n列，从0开始
	 * @return
	 */
	protected CellEditor add(int row, int col) {
		HSSFCell cell = getCell(row, col);
		workingCell.add(cell);
		return this;
	}
	
	/**
	 * 添加其他单元格到编辑队列，该方法应该在改变单元格属性前调用
	 *	否则所做的修改不会影响到新加入的单元格
	 * 
	 * @param row ExcelRowEditor
	 * @param col 第n列，从0开始
	 * @return
	 */
	protected CellEditor add(RowEditor row, int col) {
		HSSFCell cell = getCell(row.getHSSFRow(), col);
		workingCell.add(cell);
		return this;
	}
	
	/**
	 * 添加其他单元格到编辑队列，该方法应该在改变单元格属性前调用
	 *	否则所做的修改不会影响到新加入的单元格
	 * 
	 * @param row ExcelRowEditor
	 * @param col 第n列，从0开始
	 * @return
	 */
	protected CellEditor add(int row, ColumnEditor col) {
		return add(row, col.getCol());
	}
	
	/**
	 * 添加其他单元格到编辑队列，该方法应该在改变单元格属性前调用
	 *	否则所做的修改不会影响到新加入的单元格
	 * 
	 * @param row ExcelRowEditor
	 * @param col 第n列，从0开始
	 * @return
	 */
	protected CellEditor add(CellEditor cell) {
		workingCell.addAll(cell.getWorkingCell());
		return this;
	}

	/**
	 * 设置边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式
	 * @param borderColor
	 *            颜色
	 * @return
	 */
	public CellEditor border(BorderStyle borderStyle, Color borderColor) {
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setBorderBottom(borderStyle.getBorderType());
			tempCellStyle.setBorderTop(borderStyle.getBorderType());
			tempCellStyle.setBorderLeft( borderStyle.getBorderType());
			tempCellStyle.setBorderRight( borderStyle.getBorderType());
			tempCellStyle.setBottomBorderColor(borderColor.getIndex());
			tempCellStyle.setTopBorderColor(borderColor.getIndex());
			tempCellStyle.setLeftBorderColor(borderColor.getIndex());
			tempCellStyle.setRightBorderColor(borderColor.getIndex());
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置左边框
	 * @param borderStyle 样式
	 * @param borderColor 颜色
	 * @return
	 */
	public CellEditor borderLeft(BorderStyle borderStyle, Color borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderLeft(borderStyle.getBorderType());
			tempCellStyle.setLeftBorderColor(borderColor.getIndex());
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置右边框
	 * @param borderStyle 样式
	 * @param borderColor 颜色
	 * @return
	 */
	public CellEditor borderRight(BorderStyle borderStyle, Color borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderRight(borderStyle.getBorderType());
			tempCellStyle.setRightBorderColor(borderColor.getIndex());
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置上边框
	 * @param borderStyle 样式
	 * @param borderColor 颜色
	 * @return
	 */
	public CellEditor borderTop(BorderStyle borderStyle, Color borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderTop(borderStyle.getBorderType());
			tempCellStyle.setTopBorderColor(borderColor.getIndex());
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置下边框
	 * @param borderStyle 样式
	 * @param borderColor 颜色
	 * @return
	 */
	public CellEditor borderBottom(BorderStyle borderStyle, Color borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderBottom(borderStyle.getBorderType());
			tempCellStyle.setBottomBorderColor(borderColor.getIndex());
			updateCellStyle(cell);
		}
		return this;
	}

	/**
	 * 设置字体
	 * 
	 * @param fontEditor
	 *            实现IFontEditor接口
	 * @return
	 */
	public CellEditor font(IFontEditor fontEditor) {
		Map<Integer, HSSFFont> fontCache = ctx.getFontCache();
		for (HSSFCell cell : workingCell) {
			//System.out.println("===============================================");
			//System.out.println("设置单元格字体："+(cell.getCellType()== 1 ? cell.getRichStringCellValue():cell.getNumericCellValue()));
			HSSFFont font = cell.getCellStyle().getFont(workBook);
			copyFont(font, tempFont);
			fontEditor.updateFont(new Font(tempFont));
			int fontHash = tempFont.hashCode() - tempFont.getIndex();
			//System.out.println("修改字体后，计算Hash:"+fontHash);
			//System.out.println("设置的字体:"+tempFont);
			tempCellStyle.cloneStyleFrom(cell.getCellStyle());
			if (fontCache.containsKey(fontHash)) {
				//System.out.println("缓存里找到字体");
				//System.out.println("找到的字体:"+fontCache.get(fontHash)+", fontIndex="+fontCache.get(fontHash).getIndex());
				tempCellStyle.setFont(fontCache.get(fontHash));
			} else {
				//System.out.println("没找到字体，新建一个");
				HSSFFont newFont = workBook.createFont();
				copyFont(tempFont, newFont);
				//System.out.println("设置的字体:"+newFont.toString()+", fontIndex="+newFont.getIndex());
				tempCellStyle.setFont(newFont);
				int newFontHash = newFont.hashCode() - newFont.getIndex();
				fontCache.put(newFontHash, newFont);
			}
			updateCellStyle(cell);
		}
		return this;
	}

	/**
	 * 设置背景色
	 * 
	 * @param bg
	 *            颜色
	 * @return
	 */
	public CellEditor bgColor(Color bg) {
		return bgColor(bg, FillPattern.SOLID_FOREGROUND);
	}

	/**
	 * 设置背景色，可设置填充的样式
	 * 
	 * @param bg
	 *            颜色
	 * @param fillPattern 填充样式
	 * @return
	 */
	public CellEditor bgColor(Color bg, FillPattern fillPattern) {
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setFillPattern(fillPattern.getFillPattern());
			tempCellStyle.setFillForegroundColor(bg.getIndex());
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置单元格水平对齐方式
	 * @param align 对齐方式
	 * @return
	 */
	public CellEditor align(Align align){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setAlignment(align.getAlignment());
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置单元格垂直对齐方式
	 * @param align 对齐方式
	 * @return
	 */
	public CellEditor vAlign(VAlign align){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setVerticalAlignment(align.getAlignment());
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 是否自动换行。
	 * @param autoWarp true自动换行，false不换行
	 * @return
	 */
	public CellEditor warpText(boolean autoWarp){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setWrapText(autoWarp);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 是否隐藏公式，只有给工作表设置密码后，该设置才生效
	 * @param hidden true隐藏，false显示
	 * @return
	 */
	public CellEditor hidden(boolean hidden){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setHidden(hidden);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置缩进
	 * @param indent
	 * @return
	 */
	public CellEditor indent(int indent){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setIndention((short)indent);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 是否锁定，锁定后用户将不能编辑该单元格。只有给工作表设置密码后，该设置才生效
	 * @param locked true锁定，false解锁
	 * @return
	 */
	public CellEditor lock(boolean locked){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setLocked(locked);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置文字旋转角度
	 * @param rotation 角度，-90度至90度
	 * @return
	 */
	public CellEditor rotate(int rotation){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setRotation((short)rotation);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 添加批注
	 * @param content 批注内容
	 * @return
	 */
	public CellEditor comment(String content){
		HSSFPatriarch patr = ctx.getHSSFPatriarch(workingSheet);
		for (HSSFCell cell : workingCell) {
			HSSFComment comment = patr.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short)cell.getColumnIndex(), cell.getRowIndex(), (short) (cell.getColumnIndex() + 3), cell.getRowIndex() + 4));
			comment.setString(new HSSFRichTextString(content));
			cell.setCellComment(comment);
		}
		return this;
	}
	
	/**
	 * 设置自定义的样式
	 * @param style
	 * @return
	 */
	public CellEditor style(HSSFCellStyle style){
		for (HSSFCell cell : workingCell) {
			cell.setCellStyle(style);
		}
		return this;
	}
	
	/**
	 * 设置数据格式
	 * @param format 格式字符串，如0.00%，0.00E+00，#,##0等，详情请查询excel单元格格式
	 * @return
	 */
	public CellEditor dataFormat(String format){
		short index = HSSFDataFormat.getBuiltinFormat(format);
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			if(index == -1){
				HSSFDataFormat dataFormat = ctx.getWorkBook().createDataFormat();
				index = dataFormat.getFormat(format);
			}
			tempCellStyle.setDataFormat(index);
			updateCellStyle(cell);
		}		
		return this;
	}
	
	/**
	 * 设置单元格所在列的宽度
	 * @param width 宽，1表示一个字符好宽度的1/256
	 * @return
	 */
	public CellEditor width(int width){
		for (HSSFCell cell : workingCell) {
			workingSheet.setColumnWidth(cell.getColumnIndex(), width);
		}
		return this;
	}
	
	/**
	 * 增加单元格所在列的宽度
	 * @param width 要增加的宽度，1表示一个字符好宽度的1/256
	 * @return
	 */
	public CellEditor addWidth(int width){
		for (HSSFCell cell : workingCell) {
			int w = workingSheet.getColumnWidth(cell.getColumnIndex());
			workingSheet.setColumnWidth(cell.getColumnIndex(), width + w);
		}
		return this;
	}
	
	/**
	 * 设置单元格所在行的高度
	 * @param height 高，单位是像素
	 * @return
	 */
	public CellEditor height(float height){
		for (HSSFCell cell : workingCell) {
			HSSFRow row = getRow(cell.getRowIndex());
			row.setHeightInPoints(height);
		}
		return this;
	}
	
	/**
	 * 增加单元格所在行的高度
	 * @param height 要增加的高度，单位是像素
	 * @return
	 */
	public CellEditor addHeight(float height){
		for (HSSFCell cell : workingCell) {
			HSSFRow row = getRow(cell.getRowIndex());
			float h = row.getHeightInPoints();
			row.setHeightInPoints(height + h);
		}
		return this;
	}
	
	/**
	 * 获取单元格所在的行
	 * @return
	 */
	public RowEditor row(){
		return new RowEditor(workingCell.get(0).getRowIndex(), ctx);
	}
	
	/**
	 * 获取单元格所在的列
	 * @return
	 */
	public ColumnEditor colunm(){
		return new ColumnEditor(workingCell.get(0).getColumnIndex(), ctx);
	}
	
	/**
	 * 获取单元格所在的表单
	 * @return
	 */
	public SheetEditor sheet(){
		return new SheetEditor(workingCell.get(0).getSheet(), ctx);
	}

	/**
	 * 更新单元格的样式
	 * @param cell
	 */
	private void updateCellStyle(HSSFCell cell){
		Map<Integer, HSSFCellStyle> styleCache = ctx.getStyleCache();
		int tempStyleHash = tempCellStyle.hashCode() - tempCellStyle.getIndex();
		//System.out.println("寻找样式:"+tempStyleHash);
		if (styleCache.containsKey(tempStyleHash)) {
			//System.out.println("在缓存里找到样式");
			cell.setCellStyle(styleCache.get(tempStyleHash));
		} else {
			HSSFCellStyle newStyle = workBook.createCellStyle();
			newStyle.cloneStyleFrom(tempCellStyle);
			cell.setCellStyle(newStyle);
			int newStyleHash = newStyle.hashCode() - newStyle.getIndex();
			//System.out.println("新建样式，Hash="+newStyleHash);
			styleCache.put(newStyleHash, newStyle);
		}
	}
	

	/**
	 * 复制字体属性
	 * 
	 * @param src
	 *            源
	 * @param dest
	 *            目标
	 */
	private void copyFont(HSSFFont src, HSSFFont dest) {
		dest.setBoldweight(src.getBoldweight());
		dest.setCharSet(src.getCharSet());
		dest.setColor(src.getColor());
		dest.setFontHeight(src.getFontHeight());
		dest.setFontHeightInPoints(src.getFontHeightInPoints());
		dest.setFontName(src.getFontName());
		dest.setItalic(src.getItalic());
		dest.setStrikeout(src.getStrikeout());
		dest.setTypeOffset(src.getTypeOffset());
		dest.setUnderline(src.getUnderline());
	}
	
	/**
	 * 设置单元格的值
	 * 
	 * @param cell
	 *            单元格对象
	 * @param value
	 *            值
	 */
	private void setCellValue(HSSFCell cell, Object value, String pattern) {
		if (value instanceof Double || value instanceof Float || value instanceof Long || value instanceof Integer
				|| value instanceof Short || value instanceof BigDecimal) {			
			cell.setCellValue(null2Double(value.toString()));
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);//这应该在setValue之后
		} else {
			if (value != null && value.toString().startsWith("=")) {
				cell.setCellFormula(value.toString().substring(1));
				cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
			} else {
				if(value instanceof Date){//日期
					if(pattern == null || pattern.trim().equals("")){
						pattern = ctx.getDefaultStyle().getDefaultDatePattern();
					}
					SimpleDateFormat dateFormat = new SimpleDateFormat(pattern);
					cell.setCellValue(new HSSFRichTextString(dateFormat.format(value)));
				}else{
					cell.setCellValue(new HSSFRichTextString(value == null ? "" : value.toString()));
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				}
			}
		}
		invokeListener(cell, value);
	}
	
	/**
	 * 调用监听器
	 * @param cell
	 * @param value
	 */
	private void invokeListener(HSSFCell cell, Object value) {
		int sheetIndex = workBook.getSheetIndex(cell.getSheet());
		List<CellValueListener> listeners = ctx.getListenerList(sheetIndex);
		for (CellValueListener l : listeners) {
			l.onValueChange(this, value, cell.getRowIndex(),
					cell.getColumnIndex(), sheetIndex, ctx.getExcel());
		}
	}
	
	protected List<HSSFCell> getWorkingCell() {
		return workingCell;
	}
	
	/**
	 * 转换成浮点数
	 * 
	 * @param s
	 * @return
	 */
	private double null2Double(Object s) {
		double v = 0;
		if (s != null) {
			try {
				v = Double.parseDouble(s.toString());
			} catch (Exception e) {
			}
		}
		return v;
	}
}
