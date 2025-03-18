package Shopping;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class resSwiggy {
    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Swiggy Data");
        
       
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Restaurant Name");
        headerRow.createCell(1).setCellValue("Rating");
        headerRow.createCell(2).setCellValue("Delivery Time");
        headerRow.createCell(3).setCellValue("Dish Name");
        headerRow.createCell(4).setCellValue("Price");

        try {
            driver.get("https://www.swiggy.com");
            driver.manage().window().maximize();

            try {
                WebElement searchBar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/header/div/div/ul/li[5]/div/a")));
                searchBar.click();
            } catch (Exception e) {
                WebElement searchBar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/div/div/div[2]/div[2]/div[2]/div/div[1]/div")));
                searchBar.click();
            }

            WebElement searchBar1 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/div/div[1]/div/form/div/div[1]/input")));
            searchBar1.click();
            Thread.sleep(2000);

            searchBar1.sendKeys("aloo samosa");
            WebElement searchresult = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/div/div[2]/div/div/button[1]")));
            searchresult.click();

            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='Search_widgetsV2__27BBR Search_widgets__3o_bA']")));

            // Get all restaurants
            List<WebElement> restaurants = driver.findElements(By.xpath("//div[@class='Search_widgetsV2__27BBR Search_widgets__3o_bA']"));
            System.out.println(restaurants.size());

           
            int rowCount = 1; 
            for (WebElement restaurant : restaurants) {
                String restaurantName = restaurant.findElement(By.xpath(".//div[contains(@class, 'styles_restaurantName__5VIQZ styles_restaurantNameBold__2OmFY')]")).getText();
                String rating = restaurant.findElement(By.xpath(".//span[contains(@class, 'styles_restaurantMetaRating__3MhTg')]")).getText();
                String timeForDelivery = restaurant.findElement(By.xpath(".//div[contains(@class, 'styles_restaurantMeta__2yx7V')]//div[2]")).getText();
                String dishName = restaurant.findElement(By.xpath(".//div[contains(@class, 'sc-aXZVg cjJTeQ sc-hIUJlX gCYyvX')]")).getText();
                String price = restaurant.findElement(By.xpath(".//div[contains(@class, 'sc-aXZVg kCbDOU')]")).getText();

               
                System.out.println("Restaurant: " + restaurantName);
                System.out.println("Rating: " + rating);
                System.out.println("Delivery Time: " + timeForDelivery);
                System.out.println("Dish Name: " + dishName);
                System.out.println("Price: " + price);
                System.out.println("--------------------------------");

              
                Row dataRow = sheet.createRow(rowCount++);
                dataRow.createCell(0).setCellValue(restaurantName);
                dataRow.createCell(1).setCellValue(rating);
                dataRow.createCell(2).setCellValue(timeForDelivery);
                dataRow.createCell(3).setCellValue(dishName);
                dataRow.createCell(4).setCellValue(price);
            }

           
            try (FileOutputStream fileOut = new FileOutputStream(new File(".\\Output\\SwiggyData.xlsx"))) {
                workbook.write(fileOut);
            }
            System.out.println("Data written to Excel file successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
           
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            driver.quit();
        }
    }
}
