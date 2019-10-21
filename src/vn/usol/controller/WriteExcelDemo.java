package vn.usol.controller;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import vn.usol.model.Employee;
import vn.usol.model.EmployeeDAO;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class WriteExcelDemo {

	private static HSSFCellStyle createStyleForTitle(HSSFWorkbook workbook) {
		HSSFFont font = workbook.createFont();
		font.setBold(true);
		HSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);
		return style;
	}

	private static void setVal(Row row, HSSFCellStyle style, int rownum, HSSFSheet sheet) {
		Cell cell;
		// EmpNo
		cell = row.createCell(0, CellType.STRING);
		cell.setCellValue("EmpNo");
//				cell.setCellStyle(style);
		cell.setCellStyle(style);

		// EmpName
		cell = row.createCell(1, CellType.STRING);
		cell.setCellValue("EmpName");
		cell.setCellStyle(style);
		// Salary
		cell = row.createCell(2, CellType.STRING);
		cell.setCellValue("Salary");
		cell.setCellStyle(style);
		// Grade
		cell = row.createCell(3, CellType.STRING);
		cell.setCellValue("Grade");
		cell.setCellStyle(style);
		// Bonus
		cell = row.createCell(4, CellType.STRING);
		cell.setCellValue("Bonus");
		cell.setCellStyle(style);

		List<Employee> list = EmployeeDAO.listEmployees();

		// Data
		for (Employee emp : list) {
			rownum++;
			row = sheet.createRow(rownum);

			// EmpNo (A)
			cell = row.createCell(0, CellType.STRING);
			cell.setCellValue(emp.getEmpNo());
			// EmpName (B)
			cell = row.createCell(1, CellType.STRING);
			cell.setCellValue(emp.getEmpName());
			// Salary (C)
			cell = row.createCell(2, CellType.NUMERIC);
			cell.setCellValue(emp.getSalary());
			// Grade (D)
			cell = row.createCell(3, CellType.NUMERIC);
			cell.setCellValue(emp.getGrade());
			// Bonus (E)
			String formula = "0.1*C" + (rownum + 1) + "*D" + (rownum + 1);
			cell = row.createCell(4, CellType.FORMULA);
			cell.setCellFormula(formula);
		}
	}

	public static void main(String[] args) throws IOException {

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Employees sheet");
		HSSFCellStyle style = createStyleForTitle(workbook);
		HSSFCellStyle style2 = new StyleExcelDemo().getSampleStyle(workbook);

		int rownum = 0;
		Row row = sheet.createRow(rownum);
		setVal(row, style, rownum, sheet);

		File file = new File("C:/demo/employee.xls");
		file.getParentFile().mkdirs();

		FileOutputStream outFile = new FileOutputStream(file);
		workbook.write(outFile);
		System.out.println("Created file: " + file.getAbsolutePath());

	}
}