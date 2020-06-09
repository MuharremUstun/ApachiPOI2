import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ApachiPOITest {
    private final static String FILE_NAME = "src/test/resources/Books.xlsx";

    @Test
    public void test1() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = datatypeSheet.iterator();
        while (rowIterator.hasNext()) {
            Row currentRow = rowIterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                if (currentCell.getCellType() == CellType.STRING) {
                    System.out.print(currentCell.getStringCellValue() + " -- ");
                } else if (currentCell.getCellType() == CellType.NUMERIC) {
                    System.out.print(currentCell.getNumericCellValue() + " -- ");
                }
            }
            System.out.println();
        }

    }

    @Test
    public void test2() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);

        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

//        System.out.println("-----------------");

        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            int firstCell = row.getFirstCellNum();
            int lastCell = row.getLastCellNum();
            for (int j = firstCell; j < lastCell; j++) {
                System.out.print(row.getCell(j).toString() + " --");
            }
            System.out.println();
        }
    }

}
