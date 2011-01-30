package wsepr.easypoi.excel.demo;

import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor;

import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.editor.ExcelRegionEditor;
import wsepr.easypoi.excel.editor.ExcelRowEditor;
import wsepr.easypoi.excel.editor.IFontEditor;

public class CalendarDemo {
	private static final int PRINT_YEAR = 2011;
	private static final String[] days = { "日", "一", "二", "三", "四", "五", "六" };

	private static final String[] months = { "一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月",
			"九月", "十月", "十一月", "十二月" };

	public static void main(String[] args) throws Exception {
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, PRINT_YEAR);
		int year = calendar.get(Calendar.YEAR);
		Excel excel = new Excel();
		short borderColor = HSSFColor.GREY_50_PERCENT.index;
		//日期字体设置器
		IFontEditor dayFont = new IFontEditor() {
			public void updateFont(HSSFFont font) {
				font.setFontHeightInPoints((short)14);
				font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			}
		};
		//星期字体设置器
		IFontEditor monthFont = new IFontEditor() {
			public void updateFont(HSSFFont font) {
				font.setFontHeightInPoints((short)14);
				font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				font.setColor(HSSFColor.WHITE.index);
			}
		};
		for (int month = 0; month < 12; month++) {
			calendar.set(Calendar.MONTH, month);
            calendar.set(Calendar.DAY_OF_MONTH, 1);
            excel.setWorkingSheet(month)//设置第month个工作表为工作状态
            	.sheetName(months[month])//修改表名
            	.fitToPage(true)
            	.horizontallyCenter(true)
            	.displayGridlines(false);
            excel.row(0).height(80);
            //标题
            excel.cell(0, 0).value(months[month] + " " + year)
            	.align(HSSFCellStyle.ALIGN_CENTER)
            	.font(new IFontEditor() {
            		//也可以这样设置字体
					public void updateFont(HSSFFont font) {
						font.setFontHeightInPoints((short)48);
				        font.setColor(HSSFColor.DARK_BLUE.index);
					}
				});
            excel.region(0, 0, 0, 13).merge();//合并标题的单元格
            //设置星期
            for (int i = 0; i < days.length; i++) {
            	excel.column(i*2).width(5*256);
            	excel.column(i*2+1).width(13*256);
            	excel.region(1, i*2, 1, i*2+1).merge();
            	excel.cell(1, i*2).value(days[i])
        			.align(HSSFCellStyle.ALIGN_CENTER)
        			.bgColor(HSSFColor.DARK_BLUE.index)
        			.font(monthFont);
            }
            
            //开始输出日期
            int cnt = 1, day=1;
            int rownum = 2;
            for (int j = 0; j < 6; j++) {
            	ExcelRowEditor row = excel.row(rownum).height(100);//设置行高度，并返回行编辑器
                for (int i = 0; i < days.length; i++) {
                	ExcelRegionEditor dayCell = excel.region(rownum, i*2, rownum, i*2+1);
                	dayCell.align(HSSFCellStyle.ALIGN_LEFT)//设置区域内所有单元格水平对齐方式
                		.vAlign(HSSFCellStyle.VERTICAL_TOP)//设置垂直对齐方式
                		.borderOuter(HSSFCellStyle.BORDER_THIN, borderColor)//设置外边框
                		.font(dayFont);//设置字体
                    int day_of_week = calendar.get(Calendar.DAY_OF_WEEK);
                    if(cnt >= day_of_week && calendar.get(Calendar.MONTH) == month) {
                    	row.cell(i*2).value(day);//写入日期
                        calendar.set(Calendar.DAY_OF_MONTH, ++day);
                        //下面设置背景色
                        if(i == 0 || i == days.length-1) {                        	
                        	dayCell.bgColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);//周末的颜色
                        } else {
                        	dayCell.bgColor(HSSFColor.WHITE.index);//非周末
                        }
                    } else {
                    	dayCell.bgColor(HSSFColor.GREY_25_PERCENT.index);//没日期的格子
                    }
                    cnt++;
                }
                rownum++;
                if(calendar.get(Calendar.MONTH) > month) break;
            }
		}
		excel.saveExcel("F:/temp/excel/calendar.xls");//保存文件
	}
}