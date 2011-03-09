package wsepr.easypoi.excel.test;

import java.text.SimpleDateFormat;
import java.util.Date;

import wsepr.easypoi.excel.DefaultExcelStyle;
import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.style.Align;
import wsepr.easypoi.excel.style.BorderStyle;
import wsepr.easypoi.excel.style.Color;

public class ExcelTest {
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
		String excelFile = "E:/" + format.format(new Date()) + ".xls";
		DefaultExcelStyle style = new DefaultExcelStyle();
		//style.setBackgroundColor(Color.BLUE_GREY);
		//style.setFillPattern(FillPattern.SQUARES);
		style.setAlign(Align.CENTER);
		style.setFontColor(Color.CORAL);
		style.setFontSize(10);
		style.setBorderColor(Color.BLACK);
		style.setBorderStyle(BorderStyle.THIN);
		Excel excel = new Excel();
		excel.cell(0, 0).value(0.5781358).dataFormat("0.00%");
		excel.cell(0, 1).value(56489643489L).dataFormat("0.00E+00");
		excel.saveExcel(excelFile);
	}
}
