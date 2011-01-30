package wsepr.easypoi.excel;

import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import wsepr.easypoi.excel.editor.ExcelCellEditor;
import wsepr.easypoi.excel.editor.ExcelColumnEditor;
import wsepr.easypoi.excel.editor.ExcelRegionEditor;
import wsepr.easypoi.excel.editor.ExcelRowEditor;
import wsepr.easypoi.excel.editor.ExcelSheetEditor;
import wsepr.easypoi.excel.util.ExcelUtil;


/**
 * 建议使用POI 3.2以上
 * 
 * @author luxiaopan
 * 
 */
public class Excel {
	
	private ExcelContext ctx;
	
	/**
	 * 新建一个Excel文件
	 */
	public Excel() {
		this(new DefaultExcelStyle());
	}
	
	/**
	 * 新建一个Excel文件
	 * @param defaultStyle 默认样式
	 */
	public Excel(DefaultExcelStyle defaultStyle) {
		this(null, defaultStyle);
	}

	/**
	 * 用一个Excel文件作为模板创建一个Excel
	 * 
	 * @param excelPath
	 *            模板文件路径，如果模板文件不存在则创建一个空Excel文件
	 */
	public Excel(String excelPath) {
		this(excelPath, new DefaultExcelStyle());
	}
	
	/**
	 * 用一个Excel文件作为模板创建一个Excel
	 * 
	 * @param excelPath
	 *            模板文件路径，如果模板文件不存在则创建一个空Excel文件
	 * @param defaultStyle 默认样式
	 */
	public Excel(String excelPath, DefaultExcelStyle defaultStyle) {
		HSSFWorkbook workBook;
		HSSFCellStyle tempCellStyle;// 临时的样式
		HSSFFont tempFont;// 临时的字体
		
		if(excelPath == null || excelPath.trim().equals("")){
			workBook = new HSSFWorkbook();
		}else{
			workBook = readExcel(excelPath);
			if (workBook == null) {
				workBook = new HSSFWorkbook();
			}
		}
		ctx = new ExcelContext(workBook);
		ctx.setDefaultStyle(defaultStyle);
		//
		setWorkingSheet(0);
		tempCellStyle = workBook.createCellStyle();
		ctx.setTempCellStyle(tempCellStyle);
		tempFont = workBook.createFont();
		ctx.setTempFont(tempFont);
		//设置默认样式
		HSSFCell cell = ExcelUtil.getHSSFCell(ctx.getWorkingSheet(), 0, 0);
		HSSFCellStyle cellStyle = cell.getCellStyle();
		cellStyle.setFillForegroundColor(defaultStyle.getBackgroundColor());
		cellStyle.setFillPattern(defaultStyle.getFillPattern());
		cellStyle.setAlignment(defaultStyle.getAlign());
		cellStyle.setVerticalAlignment(defaultStyle.getVAlign());
		//默认字体
		HSSFFont font = cellStyle.getFont(workBook);
		font.setFontHeightInPoints(defaultStyle.getFontSize());
		font.setFontName(defaultStyle.getFontName());
		font.setColor(defaultStyle.getFontColor());
	}

	
	/**
	 * 读取模板
	 * 
	 * @param templatePath
	 *            模板文件路径
	 * @return HSSFWorkbook 返回Excel工作簿对象
	 */
	private HSSFWorkbook readExcel(String excelPath) {
		HSSFWorkbook result = null;
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(excelPath));
			result = new HSSFWorkbook(fs);
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return result;
	}

	/**
	 * 保存Excel文件
	 * 
	 * @param excelPath
	 *            保存路径
	 * @return true保存成功，false失败
	 */
	public boolean saveExcel(String excelPath) {
		boolean result = false;
		BufferedOutputStream fileOut = null;
		try {
			fileOut = new BufferedOutputStream(new FileOutputStream(excelPath));
			ctx.getWorkBook().write(fileOut);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fileOut.flush();
				fileOut.close();
			} catch (Exception e) {
				result = false;
			}
		}
		return result;
	}

	/**
	 * 选择第n个工作簿为工作状态，如果不存在则新建一个
	 * 
	 * @param index
	 *            工作簿序号，从0开始
	 * @return 第n个工作表
	 */
	public ExcelSheetEditor setWorkingSheet(int index) {
		if (index < 0) {
			index = 0;
		}
		ctx.setWorkingSheet(ExcelUtil.getHSSFSheet(ctx.getWorkBook(), index));
		return this.sheet(index);
	}

	/**
	 * 选择一个单元格
	 * 
	 * @param row
	 * @param col
	 * @return
	 */
	public ExcelCellEditor cell(int row, int col) {
		ExcelCellEditor cellEditor = new ExcelCellEditor(row, col, ctx);
		return cellEditor;
	}

	/**
	 * 选择一行
	 * @param row
	 * @return
	 */
	public ExcelRowEditor row(int row){
		ExcelRowEditor rowEditor = new ExcelRowEditor(row, ctx);
		return rowEditor;
	}
	
	/**
	 * 选择一列
	 * @param col
	 * @return
	 */
	public ExcelColumnEditor column(int col){
		ExcelColumnEditor columnEditor = new ExcelColumnEditor(col, ctx);
		return columnEditor;
	}
	
	/**
	 * 选择一个区域
	 * @param beginRow 开始行
	 * @param beginCol	开始列
	 * @param endRow 结束行
	 * @param endCol 结束列
	 * @return
	 */
	public ExcelRegionEditor region(int beginRow, int beginCol, int endRow, int endCol){
		ExcelRegionEditor regionEditor = new ExcelRegionEditor(beginRow, beginCol, endRow, endCol, ctx);
		return regionEditor;
	}
	
	/**
	 * 选择一个工作表
	 * @param index
	 * @return
	 */
	public ExcelSheetEditor sheet(int index){
		if(index < 0){
			index = 0;
		}
		ExcelSheetEditor sheetEditor = new ExcelSheetEditor(getHSSFSheet(index), ctx);
		return sheetEditor;
	}
	
	/**
	 * 选择处于工作状态的工作表
	 * @return
	 */
	public ExcelSheetEditor sheet() {
		return this.sheet(ctx.getWorkingSheetIndex());
	}
	
	/**
	 * 获取工作簿
	 * @return
	 */
	public HSSFWorkbook getWorkBook() {
		return ctx.getWorkBook();
	}

	/**
	 * 获取POI的表对象
	 * @param index 工作表序号，从0开始
	 * @return
	 */
	public HSSFSheet getHSSFSheet(int index){
		return ExcelUtil.getHSSFSheet(ctx.getWorkBook(), index);
	}
	
	/**
	 * 获取POI的行对象
	 * 
	 * @param index 工作表序号，从0开始
	 * @param row
	 *            行，从0开始
	 * @return 指定行的对象
	 */
	public HSSFRow getHSSFRow(int index, int row) {
		HSSFSheet sheet = getHSSFSheet(index);
		return ExcelUtil.getHSSFRow(sheet, row);
	}
	
	/**
	 * 获取POI的单元格对象
	 * @param index 工作表序号，从0开始
	 * @param row
	 *            行，从0开始
	 * @param col
	 *            列，从0开始
	 * @return row行col列的单元格对象
	 */
	protected HSSFCell getHSSFCell(int index, int row, int col) {
		HSSFSheet sheet = getHSSFSheet(index);
		return ExcelUtil.getHSSFCell(sheet, row, col);
	}
	
	public int getWorkingSheetIndex() {
		return ctx.getWorkingSheetIndex();
	}
}
