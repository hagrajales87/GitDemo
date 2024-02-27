import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Iterator;

public class UploadDownload {

    public WebDriver driver;
    public static FileInputStream file;

    @BeforeClass
    public void setUp(){

        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");
        driver.manage().window().setSize(new Dimension(1440, 900));
    }

    DataFormatter formatter = new DataFormatter();
    @Test
    public void testUploadDownload() throws IOException, InterruptedException {

        String fruitName = "Apple";
        int price = 983;

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(7));

        //Download
        driver.findElement(By.id("downloadButton")).click();
        Thread.sleep(1000L);
        //Edit Excel
        String fileName = "C:/Users/HectorGrajales/Downloads/download.xlsx";
        int col = getColumnName(fileName,"Price");
        int row = getRowNumber(fileName, "Apple");
        Assert.assertTrue(updateCell(fileName, row, col, price));
        System.out.println(row);
        System.out.println("Column value: " + col);


        //Upload
        WebElement uploadButton = driver.findElement(By.cssSelector("input[id='fileinput']"));
        uploadButton.sendKeys("C:/Users/HectorGrajales/Downloads/download.xlsx");


        //wait for success message to show up and wait
        By toastLocator = By.cssSelector("div[role='alert'] div:nth-child(2)");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(7));
        wait.until(ExpectedConditions.visibilityOf(driver.findElement(toastLocator)));

        Assert.assertEquals(driver.findElement(toastLocator).getText(), "Updated Excel Data Successfully.");


        wait.until(ExpectedConditions.invisibilityOfElementLocated(toastLocator));


        //verify updated excel data showing in the web table
        String priceColumn = driver.findElement(By.xpath("//div[.='Price']")).getAttribute("data-column-id");
        String actualPrice = driver.findElement(By.xpath("//div[.='"+fruitName +"']/parent::div[contains(@id,'row')]/div[@id='cell-" +priceColumn + "-undefined']")).getText();

        Assert.assertEquals(price, Integer.parseInt(actualPrice), "The expect price and current prices does not match");
        System.out.println(actualPrice);



    }

    private static boolean updateCell(String fileName, int rowNumber, int col, int value) throws IOException {
        //file = new FileInputStream(fileName);
        file = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        Row row = sheet.getRow(rowNumber);
        Cell cellField = row.getCell(col);
        cellField.setCellValue(value);
        FileOutputStream fos = new FileOutputStream(fileName);
        workbook.write(fos);
        workbook.close();
        file.close();
        return true;
    }
     private static int getRowNumber(String fileName, String fruit) throws IOException {

         file = new FileInputStream(fileName);
         XSSFWorkbook workbook = new XSSFWorkbook(file);
         XSSFSheet sheet = workbook.getSheet("Sheet1");
         boolean elementFound = false;
         Iterator<Row> rows = sheet.iterator();
         Row row = rows.next();

         int k = 0;
         int rowNumber = 0;

         while(!elementFound){
             Iterator<Cell> cell = row.cellIterator();


             while(cell.hasNext()){

                 Cell value = cell.next();
                 if(value.getCellType()== CellType.STRING){
                     if(value.getStringCellValue().equalsIgnoreCase(fruit)){
                         rowNumber = k;
                         elementFound = true;
                     }
                 }


             }
             k++;
             row = rows.next();
         }

         return rowNumber;
    }

    private static int getColumnName(String fileName, String price) throws IOException {
        file = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        Iterator<Row> rows = sheet.iterator();
        Row firstrow = rows.next();
        Iterator<Cell> cell = firstrow.cellIterator();
        Cell value = cell.next();
        int k = 0;
        int columnNumber = 0;

        while (cell.hasNext()){

            if (value.getCellType() == CellType.STRING){
                if (value.getStringCellValue().equalsIgnoreCase(price)){
                    columnNumber = k;
                }
            }
            value = cell.next();
            k++;
        }

        return columnNumber;
    }




    @AfterClass
    public void tearDown(){
        driver.quit();
    }
}
