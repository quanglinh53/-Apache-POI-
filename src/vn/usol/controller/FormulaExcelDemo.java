package vn.usol.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class FormulaExcelDemo {

	private static void getTypeFormular(Cell cell, HSSFWorkbook workbook) {
		// Formula
		String formula = cell.getCellFormula();
		
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		 
		// CellValue
		CellValue cellValue = evaluator.evaluate(cell);
		 
		double numberValue = cellValue.getNumberValue();
		String stringValue = cellValue.getStringValue();
		boolean booleanValue = cellValue.getBooleanValue();
		
		System.out.println(cellValue.getCellTypeEnum());
		System.out.println(numberValue);
		
	}

	public static void main(String[] args) throws IOException {

		File file = new File("C:/demo/employee.xls");

		// Đọc một file XSL.
		FileInputStream inputStream = new FileInputStream(file);

		// Đối tượng workbook cho file XSL.
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

		// Lấy ra sheet đầu tiên từ workbook
		HSSFSheet sheet = workbook.getSheetAt(0);

		Row row = sheet.createRow(4);
		// Create cell through type FORMULA
		Cell cell = row.createCell(2, CellType.FORMULA);
		cell.setCellFormula("SUM(C2:C4)");
		
//		Cell cell = row.createCell(2, CellType.FORMULA);
//		cell.setCellFormula("0.1*C2*D3");
		
		getTypeFormular(cell, workbook);
		
		inputStream.close();
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();

	}

}

// The process cannot access the file. Another process is in use
// Reason: File editing
