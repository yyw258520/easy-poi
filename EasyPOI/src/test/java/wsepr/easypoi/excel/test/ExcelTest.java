package wsepr.easypoi.excel.test;

import java.text.SimpleDateFormat;
import java.util.Date;

import wsepr.easypoi.excel.DefaultExcelStyle;
import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.style.Align;
import wsepr.easypoi.excel.style.BorderStyle;
import wsepr.easypoi.excel.style.Color;
import wsepr.easypoi.excel.style.FillPattern;
import wsepr.easypoi.excel.style.VAlign;

public class ExcelTest {
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
		String excelFile = "E:/" + format.format(new Date()) + ".xls";
		DefaultExcelStyle style = new DefaultExcelStyle();
		style.setBackgroundColor(Color.BLUE_GREY);
		style.setFillPattern(FillPattern.SQUARES);
		style.setAlign(Align.CENTER);
		style.setFontColor(Color.CORAL);
		style.setFontSize(20);
		Excel excel = new Excel();
		System.out.println(excel.sheet().getLastRowNum());
		excel.cell(0, 0).add(0, 1).add(0, 2).value(new Date(), "yyyy-MM-dd HH:mm").borderRight(BorderStyle.MEDIUM_DASH_DOT, Color.RED);
		excel.row(11).value(new Object[]{123123,"aabbcc",new Date(),3.1415926}).merge();
		excel.row(2).borderFull(BorderStyle.DASH_DOT, Color.BLUE_GREY).value(new Object[]{123123,"aabbcc",new Date(),3.1415926}, 2);
		excel.row(2).append(new Object[]{"添加的内容"}).borderFull(BorderStyle.SLANTED_DASH_DOT, Color.BLUE);
		excel.cell(3, 0).add(4, 0).value("合并的区间112233").rotate(90).comment("这只是一个测试的例子");
		
		excel.region(3, 3, 8, 8).image("http://www.google.com.hk/intl/zh-CN/images/logo_cn.png");
		excel.column(10).borderFull(BorderStyle.MEDIUM, Color.BLUE_GREY);
		//excel.region(4, 4, 4, 4).borderFull(HSSFCellStyle.BORDER_MEDIUM_DASH_DOT, HSSFColor.BLUE_GREY.index);
		//excel.region(15, 0, 15, 5).borderOuter(HSSFCellStyle.BORDER_DOUBLE, HSSFColor.BLUE.index);
		excel.column(10).value(new Object[]{"aaa","bbb","ccc","ddd","eee"},3).vAlign(VAlign.CENTER);
		//excel.column(0).add(1).add(2).autoWidth();
		//excel.sheet(0).sheetName("测试").header("topway", "测试的", "----");
//		excel.sheet(1).sheetName("另一个表").active();
//		excel.setWorkingSheet(1);
//		excel.cell(10, 10).add(0, 0).value("测试批注").comment("这只是一个测试");
		excel.sheet().password("123");
		excel.saveExcel(excelFile);
	}
}
