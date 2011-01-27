package wsepr.easypoi.excel.test;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import wsepr.easypoi.excel.Excel;

public class BatchExport {
	private final static int DATA_COUNT = 10000;
	private final static int FIELD_COUNT = 10;
	private final static int BATCH_SIZE = 1000;
	private static Excel excel = new Excel("E:/batch_20110125110547.xls");
	private static SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
	private static String excelFile = "E:/batch_" + format.format(new Date()) + ".xls";
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		for(int i=BATCH_SIZE;i<DATA_COUNT;i+=BATCH_SIZE){
			List<Object[]> data = initData(BATCH_SIZE);
			export(data);
		}
		excel.saveExcel(excelFile);
		try {
			Thread.sleep(10000);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}
	
	public static List<Object[]> initData(int count){
		//准备数据
		List<Object[]> data = new ArrayList<Object[]>();
		for(int i=0;i<count;i++){
			Object[] a = new Object[FIELD_COUNT];
			for(int j=0;j<FIELD_COUNT;j++){
				a[j] = "aaaaa";//RandomStringUtils.randomAlphanumeric(5);				
			}
			data.add(a);
		}
		return data;
	}
	
	private static void export(List<Object[]> data){
		int lastRow = excel.sheet().getLastRowNum();
		if(lastRow + data.size() > 65535){
			excel.setWorkingSheet(excel.getWorkingSheetIndex() + 1);
			lastRow = excel.sheet().getLastRowNum();
		}
		lastRow++;
		for(int i=0;i<data.size();i++){
			excel.row(i + lastRow).value(data.get(i));
		}
	}

}
