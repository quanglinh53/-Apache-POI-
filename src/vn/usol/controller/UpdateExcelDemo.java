package vn.usol.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class UpdateExcelDemo {

	public static void main(String[] args) throws IOException {

		File file = new File("C:/demo/employee.xls");

		// Đọc một file XSL.
		FileInputStream inputStream = new FileInputStream(file);

		// Đối tượng workbook cho file XSL.
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

		// Lấy ra sheet đầu tiên từ workbook
		HSSFSheet sheet = workbook.getSheetAt(0);

		HSSFCell cell = sheet.getRow(1).getCell(2);
		cell.setCellValue(cell.getNumericCellValue() * 2);

		cell = sheet.getRow(2).getCell(2);
		cell.setCellValue(cell.getNumericCellValue() * 2);

		cell = sheet.getRow(3).getCell(2);
		cell.setCellValue(cell.getNumericCellValue() * 2);

		inputStream.close();

		// Ghi file
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();

	}

}

// NotOLE2FileException: Invalid header signature; read 0x0000000000000000, expected 0xE11AB1A1E011CFD0 - file appears not to be a valid OLE2 document
// Reason: File not found