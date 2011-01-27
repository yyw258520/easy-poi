package wsepr.easypoi.excel.demo;

import java.util.Calendar;

import wsepr.easypoi.excel.Excel;

public class CalendarDemo {
	private static final int PRINT_YEAR = 2011;
	private static final String[] days = { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };

	private static final String[] months = { "January", "February", "March", "April", "May", "June", "July", "August",
			"September", "October", "November", "December" };

	public static void main(String[] args) throws Exception {
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, PRINT_YEAR);
		int year = calendar.get(Calendar.YEAR);
		Excel excel = new Excel();
		for (int month = 0; month < 12; month++) {
			calendar.set(Calendar.MONTH, month);
            calendar.set(Calendar.DAY_OF_MONTH, 1);
            excel.setWorkingSheet(month).sheetName(months[month]);
		}
	}
}
