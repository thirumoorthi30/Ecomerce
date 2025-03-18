package Shopping;

import org.openqa.selenium.chrome.ChromeDriver;
	
import org.testng.Assert;

import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;

import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xddf.usermodel.text.AnchorType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;


public class myntraNewThree {

    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception{
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        
        
        driver.get("https://www.amazon.in/?&tag=googhydrabk1-21&ref=pd_sl_7hz2t19t5c_e&adgrpid=155259815513&hvpone=&hvptwo=&hvadid=674842289437&hvpos=&hvnetw=g&hvrand=12817599141667678899&hvqmt=e&hvdev=c&hvdvcmdl=&hvlocint=&hvlocphy=9148906&hvtargid=kwd-10573980&hydadcr=14453_2316415&gad_source=1");
        
        driver.manage().window().maximize();

        // Find the search bar and enter the product name
        WebElement searchInput = driver.findElement(By.xpath("//*[@id=\"twotabsearchtextbox\"]"));
        searchInput.sendKeys("Trident Taupe Clover Queen Double Bedsheet Set (100% Cotton, 144 TC)");

        // Submit the search query
        //searchInput.submit();

        searchInput.sendKeys(Keys.ENTER);
        
       Thread.sleep(4000);
       
       
       try {
           // Find all product elements in the search results
           List<WebElement> productElements = driver.findElements(By.xpath("//*[@data-component-type='s-search-result']"));

           // Limit to scrape only the first three products
           int count = Math.min(productElements.size(), 3);

           // Loop through the first three product elements
           for (int i = 0; i < count; i++) {
               // Get the ith product element
               WebElement productElement = productElements.get(i);

               // Find the anchor tag within the product element
               WebElement productLink = productElement.findElement(By.xpath(".//h2/a"));

               // Get the href attribute of the anchor tag
               String productUrl = productLink.getAttribute("href");

               // Print the product URL
               System.out.println("Product URL " + (i + 1) + ": " + productUrl);
               
               
               /*       WebElement nameElement = productElement.findElement(By.xpath(".//h2/a/span"));
               String newName = nameElement.getText();
               System.out.println("Product Name " + (i + 1) + ": " + newName);
               
               
               

              String link;
               try {
            	   
              WebElement productClick = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[2]/h2/a"));
              
              
              link = productClick.getAttribute("href");
              
              System.out.println("Come on this");
              
              System.out.println(link);
              
              //Thread.sleep(1000);
               }
               catch(Exception w){
            	   WebElement productClick = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[2]/h2/a"));
            	    //  link = productClick.getText();
            	      
            	      link = productClick.getAttribute("href");
            	      System.out.println(link);
               }  
             
              
              String newName,mrp,sp,offer;
              
              Thread.sleep(15000);
              
              

            	  try {
                  	
                      WebElement nameElement = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[2]/h2/a/span"));
                      newName = nameElement.getText();
                      System.out.println(newName);
                      }
                      
                      catch(NoSuchElementException e) {
                      	
                      	WebElement nameElement = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[2]/h2/a/span"));
                      	newName = nameElement.getText();
                          System.out.println(newName);
                      	
                      }
              
            	  Thread.sleep(1000);
            	  
             
              try {
              WebElement spElement = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[4]/div[2]/div[1]/a/span/span[2]/span[2]"));
              sp = spElement.getText();
              System.out.println(sp);
              }
              
              catch(Exception t) {
            	  WebElement spElement = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[4]/div[2]/div[1]/a/span"));
                  sp = spElement.getText();
                  System.out.println(sp);
              }
              Thread.sleep(1000);
              
              //String mrp;
        	  try {
          WebElement mrpElement = driver.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[4]/div[2]/div[1]/a/div/span[2]/span[2]"));
          mrp = mrpElement.getText();
          System.out.println(mrp);
        	  }
        	  catch(Exception h) {
        		  WebElement mrpElement = driver.findElement(By.xpath(" /html/body/div[1]/div[1]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[4]/div[2]/div[1]/a/div"));
        	      mrp = mrpElement.getText();
        	      System.out.println(mrp);
        	  }
          
          Thread.sleep(1000);
        	  
              
              try {
              WebElement offerElement = driver.findElement(By.xpath("//*[@id=\"search\"]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[4]/div[2]/div[1]/span[2]"));
              offer = offerElement.getText();
              System.out.println(offer);  
              }
              
              catch(Exception n) {
            	  WebElement offerElement = driver.findElement(By.xpath("//*[@id=\"search\"]/div[1]/div[1]/div/span[1]/div[1]/div[3]/div/div/div/div/span/div/div/div[2]/div[4]/div[2]/div[1]/span[2]"));
                  offer = offerElement.getText();
                  System.out.println(offer); 
              }       */

               
               // You can extract other details of the product similarly and print them here
           }
       } catch (Exception e) {
           e.printStackTrace();
       }
       
        // Close the browser
     //   driver.quit();

    }
}