package wsepr.easypoi.excel.test;

import wsepr.easypoi.excel.Excel;

public class Test {

	/**
	 * @param args
	 * @throws NoSuchMethodException 
	 * @throws SecurityException 
	 */
	public static void main(String[] args){
		Excel excel = new Excel("F:/temp/1.xls");
//		for(int i=0;i<excel.sheet().getLastRowNum();i++){
//			System.out.println(excel.row(i));
//		}
		System.out.println(excel.column(1,1));
	}
	
	public void say(String hello, int n){
		for(int i=0;i<n;i++){
			n+= i;
		}
	}

}
