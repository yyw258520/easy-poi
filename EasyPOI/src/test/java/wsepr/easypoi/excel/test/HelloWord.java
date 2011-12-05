package wsepr.easypoi.excel.test;

import java.util.Date;

import org.apache.poi.ss.util.CellRangeAddress;

import wsepr.easypoi.excel.Excel;
import wsepr.easypoi.excel.editor.IFontEditor;
import wsepr.easypoi.excel.style.Align;
import wsepr.easypoi.excel.style.BorderStyle;
import wsepr.easypoi.excel.style.Color;
import wsepr.easypoi.excel.style.font.BoldWeight;
import wsepr.easypoi.excel.style.font.Font;

public class HelloWord {
	public static void main(String[] args) {
		Object[] val = new Object[]{"插入一行数据",123,'A',Math.PI,new Date(), "hello"};
		
		Excel excel = new Excel();
		excel.cell(0, 0) //选择第一个单元格
			.value("Hello World!")//写入值
			.align(Align.CENTER)//设置水平对齐方式
			.bgColor(Color.LIGHT_YELLOW)//设置背景色
			.height(30)//设置高度
			.font(new IFontEditor(){//设置字体
				public void updateFont(Font font) {
					font.boldweight(BoldWeight.BOLD);//粗体
					font.color(Color.BROWN);//字体颜色
				}
			});
		excel.region(0, 0, 0, 10).merge();//合并第一行10个单元格
		excel.region("$A$2:$K$2").merge();//也可以这样选取区域
		
		excel.row(2)//选择第3行
			.value(val)//写入数据
			.addWidth(2000)//增加宽度
			.borderOuter(BorderStyle.DASH_DOT_DOT, Color.CORAL);//设置外边框样式
		
		excel.row(4,1)//选择第5行，但忽略第1个单元格，从第2个单元格开始操作
			.value(val)
			.borderFull(BorderStyle.DASH_DOT, Color.RED);//设置全部边框
		
		excel.row(6)//选择第7行
			.value(val, 2)//从第3个单元格开始写入数据
			.borderTop(BorderStyle.THIN, Color.BLUE);//只设置上边框
		
		excel.column(11)
			.value(val)//也可以操作一列
			.align(Align.CENTER)
			.borderFull(BorderStyle.THICK, Color.CORNFLOWER_BLUE)//设置全部边框
			.autoWidth();//根据内容长度，自动调整列宽
		
		excel.cell(7, 0).value("=IF(B3=123,\"等于\",\"不等于\")");//写入Excel函数
		excel.cell(7, 1).value(0.578923).dataFormat("0.00%");//设置数据格式
		excel.cell(7, 2).value(0.578923, "0.00%");//也可以这样设置数据格式
		
		//插入一张图片
		excel.region(8, 0, 10, 1).image("http://poi.apache.org/resources/images/group-logo.jpg");
		
		excel.sheet().freeze(1, 0)//冻结第一行
			.sheetName("这是第一个表");//重命名当前处于工作状态的表的名称
		
		//设置单元格备注
		excel.cell(8, 5).value("这个单元格设置了备注").comment("这是一条备注");
		
		//操作第二个表
		excel.setWorkingSheet(1).sheetName("第二个表");//把第二个表设置为工作状态，并改名		
		excel.row(0).value(val);//第二个表写入数据
		excel.sheet().groupColumn(0, 3);//按列分组
		
		excel.saveExcel("E:/temp/excel/helloworld.xls");
	}
}
