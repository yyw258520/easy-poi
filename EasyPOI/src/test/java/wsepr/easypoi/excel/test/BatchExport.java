package wsepr.easypoi.excel.test;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang.RandomStringUtils;

import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.style.Align;

public class BatchExport {
	private final static int DATA_COUNT = 60000;
	private final static int FIELD_COUNT = 10;
	private final static int BATCH_SIZE = DATA_COUNT / 10;
	private static Excel excel = new Excel();
	private static SimpleDateFormat format = new SimpleDateFormat("yyyyMMddHHmmss");
	private static String excelFile = "E:/temp/batch_" + format.format(new Date()) + ".xls";
	/**
	 * @param args
	 */
	public static void main(String[] args) {
//		try {
//			Thread.sleep(10000);
//		} catch (InterruptedException e) {
//			e.printStackTrace();
//		}
		excel.row(0).value(new Object[]{"aabb","aabb","aabb","aabb","aabb"}).merge();
		excel.row(1).value(new Object[]{"aabb","aabb","aabb","aabb","这是表头"});
		long start = System.currentTimeMillis();
		for(int i=BATCH_SIZE;i<=DATA_COUNT;i+=BATCH_SIZE){
			List<Object[]> data = initData(BATCH_SIZE);
			export(data);
		}
		excel.column(0, 2).align(Align.RIGHT).value(new Object[]{1,2,3,4,5,6}).merge();	
		excel.sheet().sheetName("abcabcabcabcabcabcabcabcabcabcabcabcabcabcabcabcabcabcabcabc123455");
		excel.saveExcel(excelFile);
		long end = System.currentTimeMillis();
		System.out.println(end - start);
	}
	
	public static List<Object[]> initData(int count){
		//准备数据
		List<Object[]> data = new ArrayList<Object[]>();
		for(int i=0;i<count;i++){
			Object[] a = new Object[FIELD_COUNT];
			for(int j=0;j<FIELD_COUNT;j++){
				a[j] = RandomStringUtils.randomAlphanumeric(5);				
			}
			data.add(a);
		}
		return data;
	}
	
	private static void export(List<Object[]> data){
		excel.row();
		for(int i=0;i<data.size();i++){
			excel.row().value(data.get(i));
		}
	}

}
