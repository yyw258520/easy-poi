package wsepr.easypoi.excel.test;

import java.io.File;

import org.apache.commons.lang.RandomStringUtils;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class TestJXL {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try {
//			try {
//				Thread.sleep(10000);
//			} catch (InterruptedException e) {
//				e.printStackTrace();
//			}
			long start = System.currentTimeMillis();
			// 打开文件
			WritableWorkbook book = Workbook.createWorkbook(new File(
					"F:/temp/excel/jxl.xls"));
			// 生成名为“第一页”的工作表，参数0表示这是第一页
			WritableSheet sheet = book.createSheet(" 第一页 ", 0);
			// 在Label对象的构造子中指名单元格位置是第一列第一行(0,0)
			// 以及单元格内容为test
			for (int i = 0; i < 60000; i++) {
				for (int j = 0; j < 10; j++) {
					Label label = new Label(j, i,
							RandomStringUtils.randomAlphanumeric(5));
					sheet.addCell(label);
				}
			}
			// 写入数据并关闭文件
			book.write();
			book.close();
			long end = System.currentTimeMillis();
			System.out.println(end - start);
		} catch (Exception e) {
			System.out.println(e);
		}

	}

}
