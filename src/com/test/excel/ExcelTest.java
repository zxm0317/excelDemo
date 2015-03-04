package com.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.UUID;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("unused")
public class ExcelTest {
	List<Student> list;
	String[] headers = { "编号", "姓名", "性别", "爱好" };

	public void createExcel() {
		initData();
		int rowNumber = (list.size());
		Workbook wb = new SXSSFWorkbook(100);
		Sheet sheet = wb.createSheet();
		int cells = Student.class.getDeclaredFields().length;
		Row hrow = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++) {
			Cell c = hrow.createCell(i);
			c.setCellValue(headers[i]);
		}
		for (int i = 1; i < rowNumber; i++) {
			Student s = list.get(i - 1);
			Row row = sheet.createRow(i);
			Cell c0 = row.createCell(0);
			Cell c1 = row.createCell(1);
			Cell c2 = row.createCell(2);
			Cell c3 = row.createCell(3);
			c0.setCellValue(s.getId());
			c1.setCellValue(s.getName());
			c2.setCellValue(s.getSex());
			c3.setCellValue(s.getHobby());
		}
		try {
			FileOutputStream out = new FileOutputStream("c://sxssf.xlsx");
			wb.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void initData() {
		list = new ArrayList<Student>();
		for (int i = 0; i < 60000; i++) {
			Student s = new Student();
			s.setId(i + 1);
			s.setName("张三" + new Random().nextInt(10000));
			s.setSex("男");
			s.setHobby("打球" + UUID.randomUUID());
			list.add(s);
		}
	};

	public void readExcel() {
try {
	InputStream in=new FileInputStream(new File("c://sxssf.xlsx"));
	XSSFWorkbook wb=new XSSFWorkbook(in);
	XSSFSheet sheet = wb.getSheetAt(0);
	int rows = sheet.getPhysicalNumberOfRows();
} catch (Exception e) {
	e.printStackTrace();
}
	}

	public static void main(String[] args) {
		new ExcelTest().readExcel();
	}
}
