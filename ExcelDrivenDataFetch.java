package FrameworkData.ExcelDrivenData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDrivenDataFetch {

	public static void main(String[] args) throws IOException {
	}

	public ArrayList getData(String fieldname) throws IOException {

		ArrayList<String> ar = new ArrayList<String>();
		FileInputStream path = new FileInputStream("C:\\Users\\aman.arora\\Documents\\demodata.xlsx");

		XSSFWorkbook wb = new XSSFWorkbook(path);
		int k = 0;
		int column = 0;
		int size = wb.getNumberOfSheets();
		for (int i = 0; i < size; i++) {
			if (wb.getSheetName(i).equalsIgnoreCase("TESTDATA")) {
				XSSFSheet sh = wb.getSheetAt(i);
				Iterator<Row> rows = sh.rowIterator();
				Row firstRow = rows.next();
				Iterator<Cell> ce = firstRow.cellIterator();
				while (ce.hasNext()) {

					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {
						column = k;

					}
					k++;

				}
				System.out.println(column);
				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(fieldname)) {
						Iterator<Cell> c = r.cellIterator();
						while (c.hasNext()) {
							Cell cv = c.next();
							if (cv.getCellTypeEnum() == CellType.STRING) {
								ar.add(cv.getStringCellValue());
							} else {
								ar.add(NumberToTextConverter.toText(cv.getNumericCellValue()));
							}
						}

					}
				}

			}
		}
		return ar;
	}

}
