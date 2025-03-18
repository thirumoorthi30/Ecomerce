package DataAll;

import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ZeptoAll {

	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub

		 System.setProperty("webdriver.chrome.driver", "./Drivers/chromedriver.exe");
	        WebDriver driver = new ChromeDriver();

	        String url = "https://www.zeptonow.com/";
	        driver.get(url);
	        driver.manage().window().maximize();
	        
	        JavascriptExecutor js = (JavascriptExecutor) driver;

	     // Scroll down by 1000 pixels
	     js.executeScript("window.scrollBy(0, 2000)");
	        
	        Thread.sleep(9000);
	       
	        try {
	        	 WebElement category=driver.findElement(By.xpath("(//*[@id=\"CATEGORY_GRID_V3-element\"])[1]"));
	 	        category.click();
	        }catch (Exception e) {
	        	 WebElement category=driver.findElement(By.xpath("(//*[@id=\"CATEGORY_GRID_V3-element\"])[1]"));
		 	        category.click();
		 	        }
	        
	 //       List<WebElement> catgry=driver.findElements(By.xpath("(//div[@class='z-[20] flex'])//div[@class='plp-cat-sub-category-item flex-col plp-cat-d-view w-full md:flex-row']"));
	        
//	        for(WebElement element : catgry) {
//	    	  element.click();
//	    	  
//	    	  List<WebElement> products = driver.findElements(By.xpath("//div[@class='grid h-full w-full grid-cols-2 gap-y-4 content-start gap-x-2 px-2.5 py-4 md:p-3 md:grid-cols-3 md:gap-x-3 lg:grid-cols-5 xl:grid-cols-6 pb-24 no-scrollbar']//a"));
//	    	  
//	    	  for(WebElement product : products ) {
//	    		  
//	  	    	  WebElement name=product.findElement(By.className("//h5[@class='font-subtitle text-lg tracking-wider line-clamp-2 !text-sm !font-semibold !h-8 px-1 !tracking-wide']"));
//	  	    	  
//	  	    	  WebElement  mrp=product.findElement(By.xpath("/html/body/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/a[1]/div[3]/div/p"));
//	  	    	  
//	  	    	  WebElement sp= product.findElement(By.xpath("/html/body/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/a[1]/div[3]/div/h4"));
//	  	    	  
//	  	    	  WebElement UOM = product.findElement(By.xpath("/html/body/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/a[1]/div[2]/span/h4"));
//	  	    	  
//	  	    	  System.out.println(name);
//	  	    	  System.out.println(mrp);
//	  	    	  System.out.println(sp);
//	  	    	  System.out.println(UOM);
// 
//	    	  }
//
//	        }
	     // Cast the driver to JavascriptExecutor
	        JavascriptExecutor js1 = (JavascriptExecutor) driver;

	        // Scroll to the bottom of the page
	        js1.executeScript("window.scrollTo(0, document.body.scrollHeight);");

	        
	        List<WebElement> products = driver.findElements(By.xpath("//div[@class='grid h-full w-full grid-cols-2 gap-y-4 content-start gap-x-2 px-2.5 py-4 md:p-3 md:grid-cols-3 md:gap-x-3 lg:grid-cols-5 xl:grid-cols-6 pb-24 no-scrollbar']//a"));
	    	 
	        for(WebElement product : products ) {
	    		  
	  	    	  WebElement name=product.findElement(By.xpath("//h5[@class='font-subtitle text-lg tracking-wider line-clamp-2 !text-sm !font-semibold !h-8 px-1 !tracking-wide']"));
	  	    	  
	  	    	  WebElement  mrp=product.findElement(By.xpath("/html/body/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/a[1]/div[3]/div/p"));
	  	    	  
	  	    	  WebElement sp= product.findElement(By.xpath("/html/body/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/a[1]/div[3]/div/h4"));
	  	    	  
	  	    	  WebElement UOM = product.findElement(By.xpath("/html/body/div/div/div/div/div[2]/div[2]/div[2]/div/div/div/a[1]/div[2]/span/h4"));
	  	    	  
	  	    	  System.out.println(name);
	  	    	  System.out.println(mrp);
	  	    	  System.out.println(sp);
	  	    	  System.out.println(UOM);

	    	  }
	}

}
