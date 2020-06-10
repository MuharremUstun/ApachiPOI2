import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Task1Test {
    private static final String PATH = "src/test/resources/representative.xlsx";
    @DataProvider (name = "representativeData")
    public Object[][] representativeData() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(PATH));
        Workbook wb = new XSSFWorkbook(excelFile);
        Sheet sh = wb.getSheet("data");
        int rowCount = sh.getLastRowNum() - sh.getFirstRowNum();
        Row row = sh.getRow(0);
        int columnCount = row.getLastCellNum() - row.getFirstCellNum();
        Object[][] resultData = new Object[rowCount+1][columnCount];
        for (int i = 0; i <= rowCount; i++) {
            for (int j = 0; j < columnCount; j++) {
                resultData[i][j] = sh.getRow(i).getCell(j).toString();
            }
        }
        return resultData;
    }

    @Test (dataProvider = "representativeData")
    public void test(String s1, String s2, String s3, String s4){
        System.out.print(s1 + ", " + s2 + ", " + s3 + ", " + s4);
    }
}
