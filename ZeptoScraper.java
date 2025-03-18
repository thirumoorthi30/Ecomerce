package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.util.Arrays;
import java.util.List;

public class ZeptoScraper {
    public static void main(String[] args) throws InterruptedException {
        // Set the path to the ChromeDriver executable
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

        // Initialize the Chrome driver with headless option
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");
        WebDriver driver = new ChromeDriver(options);

        // List to store product names
        List<String> products = Arrays.asList(
            "Boroplus Doodh Kesar Body Lotion (400 ml)",
            "Cetaphil Daily Advance Ultra Hydrating Moisturizing Lotion (Dermatologists Recommended)"
        );

        // Navigate to Zepto
        driver.get("https://www.zeptonow.com/");

        // Wait for search box to be visible
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        Thread.sleep(1000);
        // Find the search bar and search for each product
        WebElement searchInput = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"__next\"]/div/div[1]/header[2]/div/div[2]/a")));

        for (int i = 0; i < Math.min(2, products.size()); i++) {
            String product = products.get(i);
            // Clear the search box and input the product name
            searchInput.click(); // Ensure focus on the search input field
            searchInput.sendKeys(Keys.CONTROL + "a", Keys.DELETE); // Clear the input field
            searchInput.sendKeys(product);
            searchInput.sendKeys(Keys.ENTER);

            // Wait for search results to load
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.product-list")));

            // Find the search results again to avoid StaleElementReferenceException
            List<WebElement> searchResults = driver.findElements(By.cssSelector("div.product-list > div.product"));

            // Extract details from the search results
            for (WebElement searchResult : searchResults) {
                try {
                    String productName = searchResult.findElement(By.cssSelector("div.details > h2")).getText();
                    String mrp = searchResult.findElement(By.cssSelector("span.mrp-price")).getText();
                    String sellingPrice = searchResult.findElement(By.cssSelector("span.sale-price")).getText();
                    System.out.println("Product Name: " + productName);
                    System.out.println("MRP: " + mrp);
                    System.out.println("Selling Price: " + sellingPrice);
                    System.out.println("-------------------------");
                } catch (org.openqa.selenium.StaleElementReferenceException ex) {
                    // If the element becomes stale, re-find it
                    searchResults = driver.findElements(By.cssSelector("div.product-list > div.product"));
                }
            }
        }

        // Close the browser after scraping the products
        driver.quit();
    }
}
