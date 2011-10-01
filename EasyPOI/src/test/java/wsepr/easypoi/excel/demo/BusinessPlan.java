package wsepr.easypoi.excel.demo;

import java.text.ParseException;

import org.apache.poi.hssf.usermodel.HSSFPrintSetup;

import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.editor.CellEditor;
import wsepr.easypoi.excel.editor.IPrintSetup;
import wsepr.easypoi.excel.editor.RowEditor;
import wsepr.easypoi.excel.editor.listener.CellValueListener;
import wsepr.easypoi.excel.style.Align;
import wsepr.easypoi.excel.style.BorderStyle;
import wsepr.easypoi.excel.style.Color;

public class BusinessPlan {
	 private static final String[] titles = {
         "ID", "Project Name", "Owner", "Days", "Start", "End",
         "9-Jul","19-Jul","23-Jul","30-Jul","6-Aug","13-Aug","20-Aug","27-Aug","3-Sep","10-Sep","17-Sep"};

	 //sample data to fill the sheet.
	 private static final Object[][] data = {
	         {"1.0", "Marketing Research Tactical Plan", "J. Dow", 70, "9-Jul", null,
	             "x", "x", "x", "x", "x", "x", "x", "x", "x", "x", "x"},
	         null,
	         {"1.1", "Scope Definition Phase", "J. Dow", 10, "9-Jul", null,
	             "x", "x", null, null,  null, null, null, null, null, null, null},
	         {"1.1.1", "Define research objectives", "J. Dow", 3, "9-Jul", null,
	                 "x", null, null, null,  null, null, null, null, null, null, null},
	         {"1.1.2", "Define research requirements", "S. Jones", 7, "10-Jul", null,
	             "x", "x", null, null,  null, null, null, null, null, null, null},
	         {"1.1.3", "Determine in-house resource or hire vendor", "J. Dow", 2, "15-Jul", null,
	             "x", "x", null, null,  null, null, null, null, null, null, null},
	         null,
	         {"1.2", "Vendor Selection Phase", "J. Dow", 19, "19-Jul", null,
	             null, "x", "x", "x",  "x", null, null, null, null, null, null},
	         {"1.2.1", "Define vendor selection criteria", "J. Dow", 3, "19-Jul", null,
	             null, "x", null, null,  null, null, null, null, null, null, null},
	         {"1.2.2", "Develop vendor selection questionnaire", "S. Jones, T. Wates", 2, "22-Jul", null,
	             null, "x", "x", null,  null, null, null, null, null, null, null},
	         {"1.2.3", "Develop Statement of Work", "S. Jones", 4, "26-Jul", null,
	             null, null, "x", "x",  null, null, null, null, null, null, null},
	         {"1.2.4", "Evaluate proposal", "J. Dow, S. Jones", 4, "2-Aug", null,
	             null, null, null, "x",  "x", null, null, null, null, null, null},
	         {"1.2.5", "Select vendor", "J. Dow", 1, "6-Aug", null,
	             null, null, null, null,  "x", null, null, null, null, null, null},
	         null,
	         {"1.3", "Research Phase", "G. Lee", 47, "9-Aug", null,
	             null, null, null, null,  "x", "x", "x", "x", "x", "x", "x"},
	         {"1.3.1", "Develop market research information needs questionnaire", "G. Lee", 2, "9-Aug", null,
	             null, null, null, null,  "x", null, null, null, null, null, null},
	         {"1.3.2", "Interview marketing group for market research needs", "G. Lee", 2, "11-Aug", null,
	             null, null, null, null,  "x", "x", null, null, null, null, null},
	         {"1.3.3", "Document information needs", "G. Lee, S. Jones", 1, "13-Aug", null,
	             null, null, null, null,  null, "x", null, null, null, null, null},
	 };
	 
	/**
	 * @param args
	 * @throws ParseException 
	 */
	public static void main(String[] args) throws ParseException {
		Excel excel = new Excel();
		//设置工作表样式
		excel.sheet().sheetName("Business Plan")//表名
			.displayGridlines(false)
			.printGridlines(false)
			.fitToPage(true)
			.horizontallyCenter(true)
			.autobreaks(true)
			.addCellValueListener(new CellListener())
			.printSetup(new IPrintSetup() {
				public void setup(HSSFPrintSetup printSetup) {
					printSetup.setLandscape(true);			        
			        printSetup.setFitHeight((short)1);
			        printSetup.setFitWidth((short)1);
				}
			});
		
		excel.row(0).value(titles).height(12.75f)
			.align(Align.CENTER).bold().bgColor(Color.LIGHT_CORNFLOWER_BLUE);
		
		excel.sheet().freeze(1, 0);//冻结第一行
		int rownum = 1;
		RowEditor rowEditor = null;
        for (int i = 0; i < data.length; i++, rownum++) {
            if(data[i] == null) continue;
            int r = rownum + 1;
            String fmla = "=IF(AND(D"+r+",E"+r+"),E"+r+"+D"+r+",\"\")";//表达式
            rowEditor = excel.row(rownum);
            rowEditor.value(data[i])//写入一行数据
            	.borderFull(BorderStyle.THIN, Color.BLACK);//设置边框样式，细线黑色
            rowEditor.cell(5).value(fmla);//写入表达式
            if(i==0 || data[i-1] == null){
            	rowEditor.bold()//把整行加粗
            			.cell(1)//选择该行第2个单元格
            			.color(Color.DARK_BLUE)//设置字体颜色
            			.warpText(true);//自动换行
            	if(i==0){
            		rowEditor.cell(1).fontHeightInPoint(14);//设置字体大小，单位像素
            	}
            }
        }
        //设置前3列的宽度
        excel.row(0).width(new int[]{256*6, 256*33, 256*20});
        
        excel.column(3).align(Align.CENTER);//居中对齐
        excel.column(4,1).dataFormat("d-mmm")//设置日期格式
					   	.align(Align.RIGHT);//靠右对齐
        excel.column(5,1).bgColor(Color.GREY_25_PERCENT)//设置背景颜色，25%灰度
    					.align(Align.RIGHT)//靠右对齐
    					.dataFormat("d-mmm");//设置日期格式
        
        //分组
        excel.sheet().groupRow(4, 6).groupRow(9, 13).groupRow(16, 18);
        
		excel.saveExcel("E:/temp/excel/BusinessPlan.xls");//保存文件
	}
	
	private static class CellListener implements CellValueListener{

		public void onValueChange(CellEditor cell, Object newValue, int row,
				int col, Excel excel) {
			if(row >=1 && col>=6 && newValue != null){
				cell.bgColor(Color.BLUE).value("");//在这里设置单元格的值将不会触发监听器
			}
		}
		
	}

}
