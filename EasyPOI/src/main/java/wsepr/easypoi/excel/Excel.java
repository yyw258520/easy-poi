package wsepr.easypoi.excel;

import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import wsepr.easypoi.excel.editor.CellEditor;
import wsepr.easypoi.excel.editor.ColumnEditor;
import wsepr.easypoi.excel.editor.RegionEditor;
import wsepr.easypoi.excel.editor.RowEditor;
import wsepr.easypoi.excel.editor.SheetEditor;
import wsepr.easypoi.excel.util.ExcelUtil;


/**
 * 建议使用POI 3.2以上
 * 
 * @author luxiaopan
 * 
 */
public class Excel {
	
	private ExcelContext ctx;
	private boolean firstRowIsLastRow = true;
	
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
		cellStyle.setFillForegroundColor(defaultStyle.getBackgroundColor().getIndex());
		cellStyle.setFillPattern(defaultStyle.getFillPattern().getFillPattern());
		cellStyle.setAlignment(defaultStyle.getAlign().getAlignment());
		cellStyle.setVerticalAlignment(defaultStyle.getVAlign().getAlignment());
		//设置边框样式
		cellStyle.setBorderBottom(defaultStyle.getBorderStyle().getBorderType());
		cellStyle.setBorderLeft(defaultStyle.getBorderStyle().getBorderType());
		cellStyle.setBorderRight(defaultStyle.getBorderStyle().getBorderType());
		cellStyle.setBorderTop(defaultStyle.getBorderStyle().getBorderType());
		cellStyle.setBottomBorderColor(defaultStyle.getBorderColor().getIndex());
		cellStyle.setTopBorderColor(defaultStyle.getBorderColor().getIndex());
		cellStyle.setLeftBorderColor(defaultStyle.getBorderColor().getIndex());
		cellStyle.setRightBorderColor(defaultStyle.getBorderColor().getIndex());
		//默认字体
		HSSFFont font = cellStyle.getFont(workBook);
		font.setFontHeightInPoints(defaultStyle.getFontSize());
		font.setFontName(defaultStyle.getFontName());
		font.setColor(defaultStyle.getFontColor().getIndex());
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
	 * @throws FileNotFoundException 
	 */
	public boolean saveExcel(String excelPath) {
		BufferedOutputStream fileOut;
		try {
			fileOut = new BufferedOutputStream(new FileOutputStream(excelPath));
			return saveExcel(fileOut);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return false;
	}
	
	/**
	 * 保存Excel文件，该方法完成操作后会关闭输出流
	 * 
	 * @param excelPath
	 *            保存路径
	 * @return true保存成功，false失败
	 */
	public boolean saveExcel(OutputStream fileOut) {
		boolean result = false;
		try {
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
	public SheetEditor setWorkingSheet(int index) {
		if (index < 0) {
			index = 0;
		}
		ctx.setWorkingSheet(ExcelUtil.getHSSFSheet(ctx.getWorkBook(), index));
		return this.sheet(index);
	}

	/**
	 * 选择一个单元格
	 * 
	 * @param row 行，从0开始
	 * @param col 列，从0开始
	 * @return
	 */
	public CellEditor cell(int row, int col) {
		CellEditor cellEditor = new CellEditor(row, col, ctx);
		return cellEditor;
	}

	/**
	 * 选择一行
	 * @param row 行，从0开始
	 * @return
	 */
	public RowEditor row(int row){
		return new RowEditor(row, ctx);
	}
	
	/**
	 * 始终选择最后一个空白行，当需要循环插入n行时特别有用
	 * @return
	 */
	public RowEditor nextRow(){
		int rowNum = ExcelUtil.getLastRowNum(ctx.getWorkingSheet());
		if(!checkEmptyRow(rowNum)){
			rowNum++;
		}
		return new RowEditor(rowNum, ctx);
	}
	
	/**
	 * 检查指定的行是否空行
	 * @param rowNum
	 * @return
	 */
	private boolean checkEmptyRow(int rowNum){
		HSSFRow row = ctx.getWorkingSheet().getRow(rowNum);
		int lastCell = row != null ? row.getLastCellNum() : 2;
		return (lastCell == 1 || lastCell == -1);
	}
	
	/**
	 * 选择一列
	 * @param col 列，从0开始
	 * @return
	 */
	public ColumnEditor column(int col){
		ColumnEditor columnEditor = new ColumnEditor(col, ctx);
		return columnEditor;
	}
	
	/**
	 * 选择一个区域
	 * @param beginRow 开始行，从0开始
	 * @param beginCol	开始列，从0开始
	 * @param endRow 结束行，从0开始
	 * @param endCol 结束列，从0开始
	 * @return
	 */
	public RegionEditor region(int beginRow, int beginCol, int endRow, int endCol){
		RegionEditor regionEditor = new RegionEditor(beginRow, beginCol, endRow, endCol, ctx);
		return regionEditor;
	}
	
	/**
	 * 选择一个工作表
	 * @param index，从0开始
	 * @return
	 */
	public SheetEditor sheet(int index){
		if(index < 0){
			index = 0;
		}
		SheetEditor sheetEditor = new SheetEditor(getHSSFSheet(index), ctx);
		return sheetEditor;
	}
	
	/**
	 * 选择处于工作状态的工作表
	 * @return
	 */
	public SheetEditor sheet() {
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
	
	/**
	 * 获取处于工作状态的工作表的需要
	 * @return 工作表序号，从0开始
	 */
	public int getWorkingSheetIndex() {
		return ctx.getWorkingSheetIndex();
	}
}
