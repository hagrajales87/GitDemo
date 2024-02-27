import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class Dataprovider {

    //multiple sets of data to our tests
    //array
    // 5 Sets of data as 5 arrays from data provider to your tests
    //then your test will run 5 times with 5 separate sets of data (arrays)

    DataFormatter formatter = new DataFormatter();
    @Test(dataProvider = "driver Test")
    public void testCaseData(String greeting, String communication,String  id){
        System.out.println(greeting + communication + id);
    }

    @DataProvider(name = "driver Test")
    public Object[][] getData() throws IOException {

        FileInputStream file = new FileInputStream("src/main/resources/excelDriven.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("testData");

        int rowCount = sheet.getPhysicalNumberOfRows();
        XSSFRow row = sheet.getRow(0);

        int collCount = row.getLastCellNum();

        XSSFCell cell;

        Object[][] data = new Object[rowCount-1][collCount];


        for(int i = 0 ; i < (rowCount-1) ; i ++){
            row = sheet.getRow(i+1);
            for(int j = 0 ; j < collCount ; j++){
                cell = row.getCell(j);
                /*
                //Option 1
                // In this scenario we must change the third (id) parameter to double
                if(cell.getCellType() == CellType.STRING){
                    //System.out.println(cell.getStringCellValue());
                    data[i][j] = cell.getStringCellValue();
                }else{
                    //System.out.println(cell.getNumericCellValue());
                    data[i][j] = cell.getNumericCellValue();
                }

                 */

                data[i][j] = formatter.formatCellValue(cell);

            }

        }

        return data;
    }
}
