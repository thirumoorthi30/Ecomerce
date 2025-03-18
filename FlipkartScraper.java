package Shopping;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.util.Arrays;
import java.util.List;

public class FlipkartScraper {
    public static void main(String[] args) {
        // Set the path to the ChromeDriver executable
    	 System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

        // Initialize the Chrome driver
    	 ChromeOptions options = new ChromeOptions();
         options.addArguments("--start-maximized");
    	 
        WebDriver driver = new ChromeDriver();

        // List of product names
        List<String> products = Arrays.asList("Boroplus Doodh Kesar Body Lotion (400 ml)", "Cetaphil Daily Advance Ultra Hydrating Moisturizing Lotion (Dermatologists Recommended)");

        // Navigate to Flipkart
        driver.get("https://www.flipkart.com/");
        
        driver.manage().window().maximize();

        for (String product : products) {
            // Find the search bar and input the product name
            WebElement searchBox = driver.findElement(By.name("q"));
            searchBox.clear();
            searchBox.sendKeys(product);

            // Submit the search query
            searchBox.submit();

            // Wait for search results to load
            // You may need to add explicit waits here

            // Find the top three search results
            List<WebElement> searchResults = driver.findElements(By.cssSelector("div._1AtVbE"));

            
            System.out.println(searchResults);
            
            
            // Extract details from the top three results
            for (int i = 0; i < Math.min(3, searchResults.size()); i++) {
                WebElement result = searchResults.get(i);
                
                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("s1Q9rs")));

                
               // WebElement productName = result.findElement(By.className("s1Q9rs"));
                String name = element.getText();
                
               
                
                WebDriverWait mrpwait = new WebDriverWait(driver, Duration.ofSeconds(10));
                WebElement mrpelement = mrpwait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div._3I9_wc")));
                
                String mrp = mrpelement.getText();
                
               // String mrp = result.findElement(By.cssSelector("div._3I9_wc")).getText();
                
                WebDriverWait spwait = new WebDriverWait(driver, Duration.ofSeconds(10));
                WebElement spelement = spwait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div._3I9_wc")));
                
                String sp = spelement.getText();
                //String sellingPrice = result.findElement(By.cssSelector("div._30jeq3")).getText();
                System.out.println("Product Name: " + name);
                System.out.println("MRP: " + mrp);
                System.out.println("Selling Price: " + sp);
                System.out.println("-------------------------");
            }
        }

        // Close the browser
        driver.quit();
    }
}
