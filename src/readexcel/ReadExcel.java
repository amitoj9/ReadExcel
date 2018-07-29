package readexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String args[]) {
		try {
			ArrayList<ExcelProperties> list = new ArrayList<ExcelProperties>();
			File myFile = new File("FILE Path");
			FileInputStream fis = new FileInputStream(myFile);
			XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
			XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			Iterator<Row> rowIterator = mySheet.iterator();
			while (rowIterator.hasNext()) {
				ExcelProperties ee = new ExcelProperties();
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				int i = 1;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					if (i == 1) {
						ee.setSRI(cell.getStringCellValue());
					}
					if (i == 2) {
						ee.setMin(cell.getNumericCellValue());
					}
					if (i == 3) {
						ee.setMax(cell.getNumericCellValue());
					}
					i++;

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:

						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t");
						break;

					}
					System.out.println("");
					list.add(ee);
				}

				ArrayList<ExcelProperties> arrayList = new ArrayList<ExcelProperties>();
				double dmin = 0;
				for (ExcelProperties excelProperties : list) {
					dmin = excelProperties.getMin();
					for (double ij = dmin; ij <= excelProperties.getMax(); ij = ij + 0.1) {
						ExcelProperties excelProperties2 = new ExcelProperties();
						excelProperties2.setSRI(excelProperties.getSRI());

						double dd = ij + 0.1;

						excelProperties2.setMin(dmin);
						excelProperties2.setMax(dd);
						dmin = dd;
						arrayList.add(excelProperties2);
					}
				}
				String excelFileName = "Output File path";// name of excel file
				String sheetName = "Sheet1";// name of sheet
				XSSFWorkbook wb = new XSSFWorkbook();
				XSSFSheet sheet = wb.createSheet(sheetName);

				// iterating r number of rows
				int r = 0;
				for (ExcelProperties excelProperties : arrayList) {

					XSSFRow rows = sheet.createRow(r);

					// iterating c number of columns
					for (int c = 0; c < 3; c++) {
						XSSFCell cell = rows.createCell(c);
						if (c == 0)
							cell.setCellValue(excelProperties.getSRI());
						if (c == 1)
							cell.setCellValue(excelProperties.getMin());
						if (c == 2)
							cell.setCellValue(excelProperties.getMax());

					}
					r++;
				}
				FileOutputStream fileOut = new FileOutputStream(excelFileName);

				// write this workbook to an Outputstream.
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();

			}
		} catch (Exception e) {
			System.out.println(e);
		}

	}
}
