package wsepr.easypoi.excel.test;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.lang.math.RandomUtils;

import wsepr.easypoi.excel.DefaultExcelStyle;
import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.editor.CellEditor;
import wsepr.easypoi.excel.editor.IFontEditor;
import wsepr.easypoi.excel.editor.listener.CellValueListener;
import wsepr.easypoi.excel.style.Align;
import wsepr.easypoi.excel.style.BorderStyle;
import wsepr.easypoi.excel.style.Color;
import wsepr.easypoi.excel.style.font.Font;

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
		excel.sheet().addCellValueListener(new CellValueListener() {
			public void onValueChange(CellEditor cell, Object newValue, int row,
					int col, int sheetIndex, Excel excel) {
				if(col == 3){
					Boolean b = (Boolean)newValue;
					if(b){
						cell.font(new IFontEditor() {
							public void updateFont(Font font) {
								font.color(Color.RED);
							}
						}).row().bgColor(Color.LIGHT_YELLOW);
					}
				}
			}
		});
		
		for(int i=0;i<10;i++){
			excel.nextRow().value(new Object[]{"test",1,new Date(), RandomUtils.nextBoolean()});
		}
		excel.saveExcel(excelFile);
	}
}
