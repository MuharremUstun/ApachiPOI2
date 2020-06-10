import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class TestWithStudents {
    private final static String FILE_NAME = "src/test/resources/students.xlsx";

    @DataProvider(name = "excelData")
    public Object[][] excelData() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheet("Data");

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

    @Test(dataProvider = "excelData")
    public void test(String c1, String c2, String c3, String c4, String c5, String c6){
        System.out.print(c1 + ", " + c2 + ", " + c3+ ", " + c4 + ", " + c5 + ", " + c6);
    }
}
