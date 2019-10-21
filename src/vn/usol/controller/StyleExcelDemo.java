package vn.usol.controller;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors;

public class StyleExcelDemo {

	// Set style for exel
	public static HSSFCellStyle getSampleStyle(HSSFWorkbook workbook) {
		// Font
		HSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setItalic(true);

		// Font Height
		font.setFontHeightInPoints((short) 18);

		// Font Color
		font.setColor(IndexedColors.RED.index);

		// Style
		HSSFCellStyle style = workbook.createCellStyle();
		style.setFont(font);

		return style;
	}

}
