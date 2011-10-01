package wsepr.easypoi.excel.demo;

import org.apache.poi.hssf.usermodel.HSSFPrintSetup;

import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.editor.IFontEditor;
import wsepr.easypoi.excel.editor.IPrintSetup;
import wsepr.easypoi.excel.style.Align;
import wsepr.easypoi.excel.style.BorderStyle;
import wsepr.easypoi.excel.style.Color;
import wsepr.easypoi.excel.style.font.Font;

public class LoanCalculator {

	public static void main(String[] args) {
		Excel excel = new Excel();
		
		excel.sheet().sheetName("Loan Calculator").printGridlines(false)
				.displayGridlines(false).fitToPage(true)
				.horizontallyCenter(true).printSetup(new IPrintSetup() {
					public void setup(HSSFPrintSetup printSetup) {
						printSetup.setLandscape(true);
					}
				});
		
		createNames(excel);
		
		ItemFont itemFont = new ItemFont();
        excel.row(0).height(35).value(new Object[]{"","","Simple Loan Calculator","","","","",""})
        	.font(itemFont)
        	.borderBottom(BorderStyle.DOTTED, Color.GREY_40_PERCENT); //标题下添加虚线
        //设置宽度
        excel.row(0).width(
				new int[] { 3 * 256, 3 * 256, 
						11 * 256, 14 * 256, 
						14 * 256, 14 * 256, 14 * 256 });
        //合并标题的单元格
        excel.region("$C$1:$H$1").merge();
       
        excel.cell(2, 4).value("Enter values").align(Align.RIGHT);
        
        excel.cell(3,2).value("Loan amount").align(Align.LEFT);
        excel.cell(3,4).align(Align.RIGHT).font(itemFont)
        	.border(BorderStyle.DOTTED, Color.GREY_40_PERCENT)
        	.dataFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)")
        	.activeCell();
        
        excel.cell(4, 2).value("Annual interest rate").align(Align.LEFT);
        excel.cell(4, 4).align(Align.RIGHT).font(itemFont)
	    	.border(BorderStyle.DOTTED, Color.GREY_40_PERCENT)
	    	.dataFormat("0.000%");
        
        excel.cell(5, 2).value("Loan period in years").align(Align.LEFT);
        excel.cell(5, 4).align(Align.RIGHT).font(itemFont)
	    	.border(BorderStyle.DOTTED, Color.GREY_40_PERCENT)
	    	.dataFormat("0");

        excel.cell(6, 2).value("Start date of loan").align(Align.LEFT);
        excel.cell(6, 4).align(Align.CENTER).font(itemFont)
	    	.dataFormat("yyyy/mm/dd");
        
        excel.cell(8, 2).value("Monthly payment").align(Align.LEFT);
        excel.cell(8, 4).align(Align.RIGHT).font(itemFont)
        	.value("=IF(Values_Entered,Monthly_Payment,\"\")")
	    	.border(BorderStyle.DOTTED, Color.GREY_40_PERCENT)
	    	.bgColor(Color.GREY_25_PERCENT)
	    	.dataFormat("$##,##0.00");
        
        excel.cell(9, 2).value("Number of payments").align(Align.LEFT);
        excel.cell(9, 4).align(Align.RIGHT).font(itemFont)
        	.value("=IF(Values_Entered,Loan_Years*12,\"\")")
	    	.border(BorderStyle.DOTTED, Color.GREY_40_PERCENT)
	    	.bgColor(Color.GREY_25_PERCENT)
	    	.dataFormat("0");
        
        excel.cell(10, 2).value("Total interest").align(Align.LEFT);
        excel.cell(10, 4).align(Align.RIGHT).font(itemFont)
        	.value("=IF(Values_Entered,Total_Cost-Loan_Amount,\"\")")
	    	.border(BorderStyle.DOTTED, Color.GREY_40_PERCENT)
	    	.bgColor(Color.GREY_25_PERCENT)
	    	.dataFormat("$##,##0.00");
        
        excel.cell(11, 2).value("Total cost of loan").align(Align.LEFT);
        excel.cell(11, 4).align(Align.RIGHT).font(itemFont)
        	.value("=IF(Values_Entered,Monthly_Payment*Number_of_Payments,\"\")")
	    	.border(BorderStyle.DOTTED, Color.GREY_40_PERCENT)
	    	.bgColor(Color.GREY_25_PERCENT)
	    	.dataFormat("$##,##0.00");

		excel.saveExcel("E:/temp/excel/loan-calculator.xls");
	}
	
	//define named ranges for the inputs and formulas
    public static void createNames(Excel excel){
    	excel.createName("Interest_Rate", "'Loan Calculator'!$E$5");
    	excel.createName("Loan_Amount", "'Loan Calculator'!$E$4");
    	excel.createName("Loan_Start", "'Loan Calculator'!$E$7");
    	excel.createName("Loan_Years", "'Loan Calculator'!$E$6");
    	excel.createName("Number_of_Payments", "'Loan Calculator'!$E$10");
    	excel.createName("Monthly_Payment", "-PMT(Interest_Rate/12,Number_of_Payments,Loan_Amount)");
    	excel.createName("Total_Cost", "'Loan Calculator'!$E$12");
    	excel.createName("Total_Interest", "'Loan Calculator'!$E$11");
    	excel.createName("Values_Entered", "IF(Loan_Amount*Interest_Rate*Loan_Years*Loan_Start>0,1,0)");
    }
    
    private static class ItemFont implements IFontEditor {
		public void updateFont(Font font) {
	        font.fontHeightInPoints(14);
	        font.fontName("Trebuchet MS");
		}
	}
}
