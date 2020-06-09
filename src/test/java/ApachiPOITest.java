import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ApachiPOITest {
    private final static String FILE_NAME = "src/test/resources/Book2.xlsx";

    @Test  // This will ignore the blank cells!!
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

    @Test  // This will NOT ignore the blank cells!!
    public void test2() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);

        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();

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

    @DataProvider (name = "excelData")
    public Object[][] excelData() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);

        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        int rowCount = lastRow - firstRow + 1;

        Row row = sheet.getRow(1);
        int firstCell = row.getFirstCellNum();
        int lastCell = row.getLastCellNum();
        int cellCount = lastCell - firstCell;

        Object[][] resultData = new Object[rowCount][cellCount];
        for (int i = firstRow; i <= lastRow; i++) {
            row = sheet.getRow(i);
            for (int j = firstCell; j < lastCell; j++) {
                resultData[i][j] = row.getCell(j).toString();
            }
        }
        return resultData;
    }

    @Test (dataProvider = "excelData")
    public void test3(String c1, String c2, String c3){
        System.out.print(c1 + ", " + c2 + ", " + c3);
    }
}
