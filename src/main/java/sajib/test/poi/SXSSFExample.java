package sajib.test.poi;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class SXSSFExample {
	public static void main(String[] args) throws Exception {
		SXSSFWorkbook workbook = new SXSSFWorkbook(100);
		Sheet sheet = workbook.createSheet();
		for (int rownum = 0; rownum < 500; rownum++) {
			Row row = sheet.createRow(rownum);
			for (int cellnum = 0; cellnum < 10; cellnum++) {
				Cell cell = row.createCell(cellnum);
				cell.setCellValue((rownum + 1) + ", " + (cellnum + 1));
			}
		}

		System.out.println(sheet.getRow(0));
		System.out.println(sheet.getRow(200));
		System.out.println(sheet.getRow(400));

		final String FILE_NAME = "./sxssf_example.xlsx";
		FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
		workbook.write(outputStream);
		outputStream.close();
		workbook.dispose();
		workbook.close();
	}
}
