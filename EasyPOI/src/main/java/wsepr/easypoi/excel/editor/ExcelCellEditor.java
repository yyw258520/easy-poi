package wsepr.easypoi.excel.editor;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;

import wsepr.easypoi.excel.ExcelContext;


public class ExcelCellEditor extends AbstractExcelEditor{
	
	/**
	 * 缓存style对象
	 */
	private Map<Integer, HSSFCellStyle> styleCache = new HashMap<Integer, HSSFCellStyle>();
	private Map<Integer, HSSFFont> fontCache = new HashMap<Integer, HSSFFont>();
	private List<HSSFCell> workingCell = new ArrayList<HSSFCell>(2);

	public ExcelCellEditor(int row, int col, ExcelContext context) {
		this(context);
		this.add(row, col);
	}
	
	public ExcelCellEditor(ExcelContext context) {
		super(context);
		init();
	}

	private void init(){
		short numStyle = this.workBook.getNumCellStyles();
		for(short i=0; i<numStyle;i++){
			HSSFCellStyle style = this.workBook.getCellStyleAt(i);
			if(style != this.tempCellStyle){
				this.styleCache.put(style.hashCode() - style.getIndex(), style);
			}
		}
		short numFont = this.workBook.getNumberOfFonts();
		for(short i=0; i<numFont;i++){
			HSSFFont font = this.workBook.getFontAt(i);
			if(font != this.tempFont){
				this.fontCache.put(font.hashCode() - font.getIndex(), font);
			}
		}
	}
	
	/**
	 * 写入内容
	 * 
	 * @param value
	 * @return
	 */
	public ExcelCellEditor value(Object value) {
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
	public ExcelCellEditor value(Date value, String pattern) {
		for (HSSFCell cell : workingCell) {
			this.setCellValue(cell, value, pattern);
		}
		return this;
	}

	/**
	 * 添加其他单元格到编辑队列，该方法应该在改变单元格属性前调用
	 *	否则所做的修改不会影响到新加入的单元格
	 * 
	 * @param row
	 * @param col
	 * @return
	 */
	public ExcelCellEditor add(int row, int col) {
		HSSFCell cell = getCell(row, col);
		workingCell.add(cell);
		return this;
	}

	/**
	 * 设置边框样式和颜色
	 * 
	 * @param borderStyle
	 *            样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public ExcelCellEditor border(int borderStyle, int borderColor) {
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setBorderBottom((short) borderStyle);
			tempCellStyle.setBorderTop((short) borderStyle);
			tempCellStyle.setBorderLeft((short) borderStyle);
			tempCellStyle.setBorderRight((short) borderStyle);
			tempCellStyle.setBottomBorderColor((short) borderColor);
			tempCellStyle.setTopBorderColor((short) borderColor);
			tempCellStyle.setLeftBorderColor((short) borderColor);
			tempCellStyle.setRightBorderColor((short) borderColor);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置左边框
	 * @param borderStyle 样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor 颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public ExcelCellEditor borderLeft(int borderStyle, int borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderLeft((short) borderStyle);
			tempCellStyle.setLeftBorderColor((short) borderColor);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置右边框
	 * @param borderStyle 样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor 颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public ExcelCellEditor borderRight(int borderStyle, int borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderRight((short) borderStyle);
			tempCellStyle.setRightBorderColor((short) borderColor);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置上边框
	 * @param borderStyle 样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor 颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public ExcelCellEditor borderTop(int borderStyle, int borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderTop((short) borderStyle);
			tempCellStyle.setTopBorderColor((short) borderColor);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置下边框
	 * @param borderStyle 样式，例如HSSFCellStyle.BORDER_MEDIUM
	 * @param borderColor 颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public ExcelCellEditor borderBottom(int borderStyle, int borderColor){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setBorderBottom((short) borderStyle);
			tempCellStyle.setBottomBorderColor((short) borderColor);
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
	public ExcelCellEditor font(IFontEditor fontEditor) {
		for (HSSFCell cell : workingCell) {
			HSSFFont font = cell.getCellStyle().getFont(workBook);
			copyFont(font, tempFont);
			fontEditor.updateFont(tempFont);
			int fontHash = tempFont.hashCode() - tempFont.getIndex();
			tempCellStyle.cloneStyleFrom(cell.getCellStyle());
			if (fontCache.containsKey(fontHash)) {
				tempCellStyle.setFont(fontCache.get(fontHash));
			} else {
				HSSFFont newFont = workBook.createFont();
				copyFont(tempFont, newFont);
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
	 *            颜色，例如HSSFColor.RED.index
	 * @return
	 */
	public ExcelCellEditor bgColor(short bg) {
		return bgColor(bg, HSSFCellStyle.SOLID_FOREGROUND);
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
	public ExcelCellEditor bgColor(short bg, short fillPattern) {
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setFillPattern(fillPattern);
			tempCellStyle.setFillForegroundColor(bg);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置单元格水平对齐方式
	 * @param align 对齐方式，例如HSSFCellStyle.ALIGN_CENTER。可选值：ALIGN_GENERAL, ALIGN_LEFT, ALIGN_CENTER, ALIGN_RIGHT, ALIGN_FILL, ALIGN_JUSTIFY, ALIGN_CENTER_SELECTION
	 * @return
	 */
	public ExcelCellEditor align(Short align){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setAlignment(align);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 设置单元格垂直对齐方式
	 * @param align 对齐方式，例如HSSFCellStyle.VERTICAL_TOP。可选值：VERTICAL_TOP, VERTICAL_CENTER, VERTICAL_BOTTOM, VERTICAL_JUSTIFY
	 * @return
	 */
	public ExcelCellEditor vAlign(short align){
		for (HSSFCell cell : workingCell) {
			HSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			//
			tempCellStyle.setVerticalAlignment(align);
			updateCellStyle(cell);
		}
		return this;
	}
	
	/**
	 * 是否自动换行。
	 * @param autoWarp true自动换行，false不换行
	 * @return
	 */
	public ExcelCellEditor warpText(boolean autoWarp){
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
	 * 添加批注
	 * @param content 批注内容
	 * @return
	 */
	public ExcelCellEditor comment(String content){
		HSSFPatriarch patr = ctx.getHSSFPatriarch(this.workingSheet);
		for (HSSFCell cell : workingCell) {
			HSSFComment comment = patr.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short)cell.getColumnIndex(), cell.getRowIndex(), (short) (cell.getColumnIndex() + 3), cell.getRowIndex() + 4));
			comment.setString(new HSSFRichTextString(content));
			cell.setCellComment(comment);
		}
		return this;
	}

	/**
	 * 更新单元格的样式
	 * @param cell
	 */
	private void updateCellStyle(HSSFCell cell){
		int tempStyleHash = tempCellStyle.hashCode() - tempCellStyle.getIndex();
		if (styleCache.containsKey(tempStyleHash)) {
			cell.setCellStyle(styleCache.get(tempStyleHash));
		} else {
			HSSFCellStyle newStyle = workBook.createCellStyle();
			newStyle.cloneStyleFrom(tempCellStyle);
			cell.setCellStyle(newStyle);
			int newStyleHash = newStyle.hashCode() - newStyle.getIndex();
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
	protected void setCellValue(HSSFCell cell, Object value, String pattern) {
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
