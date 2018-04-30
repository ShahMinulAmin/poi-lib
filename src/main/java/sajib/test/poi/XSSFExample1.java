package sajib.test.poi;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFExample1 {
	public static void main(String[] args) throws Exception {	
		String[] books = { "The Tempest", "Gitanjali", "Harry Potter" };
		String[] authors = { "William Shakespeare", "Rabindranath Tagore", "J. K. Rowling" };

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		sheet.setColumnWidth((short) 0, (short) ((50 * 8) / ((double) 1 / 20)));
		sheet.setColumnWidth((short) 1, (short) ((50 * 8) / ((double) 1 / 20)));
		workbook.setSheetName(0, "XSSFWorkbook example");

		Font font1 = workbook.createFont();
		font1.setFontHeightInPoints((short) 10);
		font1.setColor((short) 0xc); // make it blue
		font1.setBold(true);
		XSSFCellStyle cellStyle1 = (XSSFCellStyle) workbook.createCellStyle();
		cellStyle1.setFont(font1);

		Font font2 = workbook.createFont();
		font2.setFontHeightInPoints((short) 10);
		font2.setColor((short) Font.COLOR_NORMAL);
		XSSFCellStyle cellStyle2 = (XSSFCellStyle) workbook.createCellStyle();
		cellStyle2.setFont(font2);

		Row headerRow = sheet.createRow(0);
		Cell cell1 = headerRow.createCell(0);
		cell1.setCellValue("Book");
		cell1.setCellStyle(cellStyle1);
		Cell cell2 = headerRow.createCell(1);
		cell2.setCellValue("Author");
		cell2.setCellStyle(cellStyle1);

		int rownum; Row row = null; Cell cell = null;
		for (rownum = (short) 1; rownum <= books.length; rownum++) {
			row = sheet.createRow(rownum);
			cell = row.createCell(0);
			cell.setCellValue(books[rownum - 1]);
			cell.setCellStyle(cellStyle2);
			cell = row.createCell(1);
			cell.setCellValue(authors[rownum - 1]);
			cell.setCellStyle(cellStyle2);
		}

		final String FILE_NAME = "./xssf_example.xlsx";
		FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
	}
}
