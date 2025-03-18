package Shopping;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.InvalidArgumentException;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class code{
    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

        WebDriver driver = new ChromeDriver();

        try {
            String filePath = ".\\input-data\\firstcry.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(file);

            Sheet sheet = workbook.getSheet("Sheet1");
            int rowCount = sheet.getPhysicalNumberOfRows();

            List<String> urls = new ArrayList<>();

            for (int i = 0; i < rowCount; i++) {
            	  Row row = sheet.getRow(i);
                Cell cell = row.getCell(0);

                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String url = cell.getStringCellValue();
                    urls.add(url);
                }
            }

          //  file.close();

            for (String url : urls) {
                try {
                	
                //	driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
                	
                    System.out.println("Processing URL: " + url);
                    if(url == " ") {
                    	driver.close();
                    }
                    driver.get(url);
                    driver.manage().window().maximize();

                    WebElement nameElement = driver.findElement(By.id("prod_name"));
                    WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));

                    String mrp1 = mrp.getText();
                    String name = nameElement.getText();

                    System.out.println(name);
                    System.out.println(mrp1);
                   
                    for (int j = 0; j < 150; j++) {
                       // driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                    	driver.findElement(By.xpath("//*[@id=\"p_breadcrumb\"]/div[2]/div/div[2]/div[6]/div/div[2]/div[2]/div[1]/div/span[1]")).click();
                		
                		driver.findElement(By.xpath("//*[@id=\"p_breadcrumb\"]/div[2]/div/div[2]/div[6]/div/div[2]/div[2]/div[1]/div/span[2]")).click();
                        
                        //driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                        break;
                    }

                    
               //     driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
                    
              /*      By elementLocator = By.className("EMINetPayment");

                    Duration maxWaitTimeInSeconds = Duration.ofSeconds(50);

                    WebDriverWait wait = new WebDriverWait(driver, maxWaitTimeInSeconds);

                    WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(elementLocator));    */

                    WebElement rate = driver.findElement(By.className("EMINetPayment"));
                    String rate1 = rate.getText();

                   System.out.println(rate1);
                   

                    for (int k = 0; k < 350; k++) {
                    	
                    	
                       
                        driver.findElement(By.xpath("//*[@id=\"delete_12022306\"]")).click();
                        break;
                    } 
                    
                    for (int  i= 0; i < rowCount; i++) {
                    	  Row row = sheet.getRow(i);
                        Cell cell = row.getCell(0);

                    Cell mrpCell = row.createCell(1);
                      Cell nameCell = row.createCell(2);
                      Cell rateCell = row.createCell(3);
                      
                      mrpCell.setCellValue(mrp1);
                      nameCell.setCellValue(name);
                      rateCell.setCellValue(rate1);  
                      
                    System.out.println(nameCell); 
                      System.out.println(mrpCell); 
                      System.out.println(rateCell); 

                      
                     // driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
                      
                   //  driver.navigate().back();
                      }
               

                }
                
                catch (Exception e) {
                    e.printStackTrace();
                }
            }

            
            
            FileOutputStream outFile = new FileOutputStream(".\\Output\\output Data.xlsx");
            workbook.write(outFile);
            outFile.close();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
               // driver.quit();
            }
        }
    }
}
