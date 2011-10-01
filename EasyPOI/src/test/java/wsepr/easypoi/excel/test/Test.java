package wsepr.easypoi.excel.test;

import java.util.Date;

import wsepr.easypoi.excel.Excel;

public class Test {

	/**
	 * @param args
	 * @throws NoSuchMethodException 
	 * @throws SecurityException 
	 */
	public static void main(String[] args){
		Excel excel = new Excel();
//		for(int i=0;i<excel.sheet().getLastRowNum();i++){
//			System.out.println(excel.row(i));
//		}
		excel.cell(0, 0).value(new Date(),"h:m:s");
		excel.cell(0, 1).value((byte)15);
		excel.cell(0, 2).value(true);
		excel.cell(0, 3).value(Math.PI,"0.00");
		excel.saveExcel("E:/temp/excel/1.xls");
	}
	
	public void say(String hello, int n){
		for(int i=0;i<n;i++){
			n+= i;
		}
	}

}
