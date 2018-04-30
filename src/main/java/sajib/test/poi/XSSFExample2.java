package sajib.test.poi;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XSSFExample2 {
	public static void main(String[] args) throws Exception {
		final String FILE_NAME = "./xssf_example.xlsx";
		FileInputStream excelInputStream = new FileInputStream(new File(FILE_NAME));
		Workbook workbook = new XSSFWorkbook(excelInputStream);
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowItr = sheet.iterator();
		int rowNum = 0;

		while (rowItr.hasNext()) {
			Row row = rowItr.next();
			Iterator<Cell> cellItr = row.iterator();
			System.out.print(rowNum + ". ");
			while (cellItr.hasNext()) {
				Cell cell = cellItr.next();
				if (cell.getCellTypeEnum() == CellType.STRING) {
					System.out.print(cell.getStringCellValue() + "\t\t");
				} else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
					System.out.print(cell.getNumericCellValue() + "\t\t");
				}
			}
			System.out.println();
			rowNum++;
		}
		workbook.close();
		excelInputStream.close();
	}
}
