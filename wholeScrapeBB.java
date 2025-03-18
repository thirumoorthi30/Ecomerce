package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.util.List;

public class wholeScrapeBB {
    public static void main(String[] args) throws InterruptedException {
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.bigbasket.com/pc/fruits-vegetables/cuts-sprouts/cut-peeled-veggies/?nc=nb");
        driver.manage().window().maximize();
        
        // Add wait here if needed
        Thread.sleep(5000);
        List<WebElement> categoryLinks = driver.findElements(By.xpath("//a[@class='h-full']"));

        for (WebElement categoryLink : categoryLinks) {
        	String productBrandName = "NA";
            String productUrl = "NA";
            String productNameResult = "NA";
            String originalMrp = "NA";
            String originalUom = "NA";
            String spValue = "NA";
            String offer = "NA";
            
            Thread.sleep(2000); // Add wait here if needed
            //products Link
            try {
                productUrl = categoryLink.getAttribute("href");
            } catch (Exception e) {
            	 e.printStackTrace();
            }
            Thread.sleep(1000);
            //Product Brand
            for (int i = 0; i < 200; i++) {
            	 try {
                     productBrandName = categoryLink.findElement(By.xpath(".//span[@class='Label-sc-15v1nk5-0 BrandName___StyledLabel2-sc-hssfrl-1 gJxZPQ keQNWn']")).getText();
                     break;
                 } catch (Exception e) {
                	 if (i==199) {
                		 e.printStackTrace();
					}
                 	
                 }
			}
           
            Thread.sleep(1000);
            //Product name
            try {
                productNameResult = categoryLink.findElement(By.xpath(".//div[@class='break-words h-10 w-full']//h3")).getText();
            } catch (Exception e) {
            	 e.printStackTrace();
            }
            
            Thread.sleep(1000);
            //mrp
            
            for(int i=0;i<200;i++) {
            try {
                originalMrp = categoryLink.findElement(By.xpath("//h3[contains(text(),'"+productNameResult+"')]/ancestor::h3/following-sibling::div//div[contains(@class,'Pricing___')]/span[2]")).getText().replace("₹", "");
                break;
            } catch (Exception e) {
            	if (i==199) {
         		   e.printStackTrace();
				}
                
            }
            }

            Thread.sleep(1000);
            //sp
            
            for(int j=0;j<200;j++) {
            	
            
            try {
                spValue = categoryLink.findElement(By.xpath("//h3[contains(text(),'"+productNameResult+"')]/ancestor::h3/following-sibling::div//div[contains(@class,'Pricing___')]/span[1]")).getText().replace("₹", "");
                break;
            } catch (Exception e) {
            	if (j==199) {
            		 e.printStackTrace();
				}
            	
            }
        }
        
           Thread.sleep(1000);
           for (int i = 0; i < 200; i++) {
        	 //uom
               try {
                   originalUom = categoryLink.findElement(By.xpath("//h3[contains(text(),'"+productNameResult+"')]/ancestor::a/following-sibling::div[2]//span")).getText();
                   break;
               } catch (Exception e) {
            	   if (i==199) {
            		   e.printStackTrace();
				}
               	
               }
		}
            
            Thread.sleep(1000);
          //uom
            for (int i = 0; i < 200; i++) {
            	try {
                    offer = categoryLink.findElement(By.xpath("//h3[contains(text(),'"+productNameResult+"')]/ancestor::h3/parent::div/div[1]//div[contains(@class,'Offers')]//span[contains(text(),'% ')]")).getText().replace("OFF", "Off");
                    break;
                } catch (Exception e) {
                	 if (i==199) {
              		   e.printStackTrace();
  				}
                }
			}
            

            // Process scraped data here
            System.out.println("Product Brand: " + productBrandName);
            System.out.println("Product URL: " + productUrl);
            System.out.println("Product Name: " + productNameResult);
            System.out.println("Original MRP: " + originalMrp);
            System.out.println("SP Value: " + spValue);
            System.out.println("Original UOM: " + originalUom);
            System.out.println("Original Offer: " + offer);
            
        }

        driver.quit();
    }
}