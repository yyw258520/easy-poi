package wsepr.easypoi.excel.test;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;

import wsepr.easypoi.excel.DefaultExcelStyle;
import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.editor.IFontEditor;

public class ExcelTest {
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
		String excelFile = "E:/" + format.format(new Date()) + ".xls";
		DefaultExcelStyle style = new DefaultExcelStyle();
		style.setBackgroundColor(HSSFColor.WHITE.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setAlign(HSSFCellStyle.ALIGN_CENTER);
		style.setFontColor(HSSFColor.RED.index);
		style.setFontSize((short)20);
		Excel excel = new Excel();
		System.out.println(excel.sheet().getLastRowNum());
		excel.cell(0, 0).add(0, 1).add(0, 2).value(new Date(), "yyyy-MM-dd HH:mm").borderRight(HSSFCellStyle.BORDER_MEDIUM, HSSFColor.RED.index);
		excel.cell(1, 0).add(1, 1)
			.border(HSSFCellStyle.BORDER_DASH_DOT_DOT, HSSFColor.BLACK.index)
			.value("测试一下测试一下测试一下测试一下测试一下测试一下")
			.bgColor(HSSFColor.BLUE.index)
			.font(new IFontEditor(){
				public void updateFont(HSSFFont font) {
					font.setFontName("黑体");
					font.setColor(HSSFFont.COLOR_NORMAL);
					font.setUnderline(HSSFFont.U_DOUBLE);
					font.setItalic(true);
					font.setFontHeightInPoints((short)18);
				}
			})
			.align(HSSFCellStyle.ALIGN_CENTER);		
		excel.row(11).value(new Object[]{123123,"aabbcc",new Date(),3.1415926}).merge();
		excel.row(2).borderFull(HSSFCellStyle.BORDER_MEDIUM_DASH_DOT, HSSFColor.BLUE_GREY.index).value(new Object[]{123123,"aabbcc",new Date(),3.1415926}, 2);
		excel.row(2).append(new Object[]{"添加的内容"}).borderFull(HSSFCellStyle.BORDER_DOUBLE, HSSFColor.BLUE.index);
		excel.cell(3, 0).add(4, 0).value("合并的区间112233").comment("这只是一个测试的例子");
		
		excel.region(3, 3, 8, 8).image("http://www.google.com.hk/intl/zh-CN/images/logo_cn.png");
		excel.column(10).borderFull(HSSFCellStyle.BORDER_MEDIUM_DASH_DOT, HSSFColor.BLUE_GREY.index);
		//excel.region(4, 4, 4, 4).borderFull(HSSFCellStyle.BORDER_MEDIUM_DASH_DOT, HSSFColor.BLUE_GREY.index);
		//excel.region(15, 0, 15, 5).borderOuter(HSSFCellStyle.BORDER_DOUBLE, HSSFColor.BLUE.index);
		excel.column(10).value(new Object[]{"aaa","bbb","ccc","ddd","eee"},3).vAlign(HSSFCellStyle.VERTICAL_CENTER);
		//excel.column(0).add(1).add(2).autoWidth();
		//excel.sheet(0).sheetName("测试").header("topway", "测试的", "----");
//		excel.sheet(1).sheetName("另一个表").active();
//		excel.setWorkingSheet(1);
//		excel.cell(10, 10).add(0, 0).value("测试批注").comment("这只是一个测试");
		excel.saveExcel(excelFile);
	}
}
