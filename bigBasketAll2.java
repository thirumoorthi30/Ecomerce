package DataAll;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class bigBasketAll2 {

    public static void main(String[] args) throws Exception {

        WebDriver driver = null;
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Product Data");
        int rowCount = 0;

        // Read input data from Excel file containing category URLs and Location Data
        FileInputStream fis = new FileInputStream(".\\input-data\\CategoryLinks.xlsx");
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        String lastLocation = "";
        boolean hasValidLinks = false;

        try {
            // Initialize the WebDriver and set up Chrome
            driver = new ChromeDriver();
            driver.manage().window().maximize();
            Thread.sleep(1000); // Allow time for browser to initialize

            // Iterate through each row of input file (category URLs and locations)
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String categoryNAME = row.getCell(0) != null ? row.getCell(0).getStringCellValue() : "";
                String categoryUrl = row.getCell(1) != null ? row.getCell(1).getStringCellValue() : "";
                String location = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "";
                String pincode = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : "";  // Assuming pincode is in 4th column

                if (categoryUrl.isEmpty() || categoryNAME.isEmpty()) {
                    System.out.println("Skipping empty row: No URL or category name provided.");
                    continue;
                }

                // Flag to indicate valid URLs were found
                hasValidLinks = true;

                // Set the location only if it's different from the last one
                if (!location.equals(lastLocation)) {
                	driver.get("https://www.bigbasket.com/");
                    Thread.sleep(1000);
                	
                    System.out.println(pincode);
                    System.out.println(location);
                    
                    
                	handleLocationSelection(driver, location, pincode);
                    lastLocation = location;
                }

                System.out.println("Processing category URL: " + categoryUrl);

                // Load the BigBasket website
                

                // Handle location selection dynamically
             //   handleLocationSelection(driver, location, pincode);

                Thread.sleep(2000);

                // Navigate to the category page
                driver.get(categoryUrl);
                Thread.sleep(1000); // Wait for the page to load

                List<WebElement> categoryLinks = new ArrayList<>();
                try {
                    // Scroll and scrape the links for the category page
                    scrollAndScrape(driver);
                    categoryLinks = scrapeCategoryLinks(driver); // Now we get the links after scrolling
                    
                    System.out.println("Total number of product in the cato is ===== " + categoryLinks.size());
                } catch (Exception e) {
                    System.out.println("Error occurred while scraping category links. Retrying...");
                    scrollAndScrape(driver);  // Retry scroll and scrape
                    categoryLinks = scrapeCategoryLinks(driver); // Retry fetching links after scroll
                    System.out.println("Total number of product in the cato is ===== " + categoryLinks.size());
                }

                System.out.println("Total number of products found: " + categoryLinks.size());

                // Scrape product details for each product link
                for (WebElement categoryLink : categoryLinks) {
                    String productBrandName = "NA";
                    String productUrl = "NA";
                    String productNameResult = "NA";
                    String originalMrp = "NA";
                    String originalUom = "NA";
                    String spValue = "NA";
                    String offer = "NA";
                    String productCode = "NA";

                    try {
                        productUrl = getProductUrl(categoryLink);
                        
                        productCode = extractProductCode(productUrl);
                    } catch (Exception e) {
                        productUrl = "NA";
                        
                        productCode = "NA";
                    }

                    try {
                        productBrandName = getProductBrandName(categoryLink);
                    } catch (Exception e) {
                        productBrandName = "NA";
                    }

                    try {
                        productNameResult = getProductName(categoryLink);
                    } catch (Exception e) {
                        productNameResult = "NA";
                    }

                    try {
                        originalMrp = getOriginalMrp(categoryLink);
                    } catch (Exception e) {
                        originalMrp = "NA";
                    }

                    try {
                        spValue = getSpValue(categoryLink);
                    } catch (Exception e) {
                        spValue = "NA";
                    }

                    try {
                        originalUom = getOriginalUom(categoryLink,productNameResult);
                    } catch (Exception e) {
                        originalUom = "NA";
                    }

                    try {
                        offer = getOffer(categoryLink,productNameResult);
                    } catch (Exception e) {
                        offer = "NA";
                    }

                    // Write the product details to the output Excel file
                    Row outputRow = outputSheet.createRow(rowCount++);
                    outputRow.createCell(0).setCellValue(productBrandName);
                    outputRow.createCell(1).setCellValue(productUrl);
                    outputRow.createCell(2).setCellValue(productNameResult);
                    outputRow.createCell(3).setCellValue(originalMrp);
                    outputRow.createCell(4).setCellValue(spValue);
                    outputRow.createCell(5).setCellValue(originalUom);
                    outputRow.createCell(6).setCellValue(offer);
                    outputRow.createCell(7).setCellValue(categoryNAME);
                    outputRow.createCell(8).setCellValue(pincode);
                    outputRow.createCell(9).setCellValue(productCode);

                    System.out.println("Product Brand: " + productBrandName);
                    System.out.println("Product URL: " + productUrl);
                    System.out.println("Product Name: " + productNameResult);
                    System.out.println("Original MRP: " + originalMrp);
                    System.out.println("SP Value: " + spValue);
                    System.out.println("Original UOM: " + originalUom);
                    System.out.println("Offer: " + offer);
                    System.out.println("Offer: " + productCode);
                }

                // Save the output after processing each location
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Bigbasket_Fully_OutputData_" + location + "_" + timestamp + ".xlsx";

                // Ensure the output directory exists before writing to the file
                File outputFileDir = new File(".\\Output");
                if (!outputFileDir.exists()) {
                    outputFileDir.mkdirs();  // Create the directory if it doesn't exist
                }

                try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
                    outputWorkbook.write(outFile);
                }
                System.out.println("Output file saved for location: " + location + " at " + outputFilePath);
            }

            // If no valid links were found, exit
            if (!hasValidLinks) {
                System.out.println("No valid links found in the input file. Exiting.");
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
             //   driver.quit();  // Close the WebDriver after execution
                System.out.println("WebDriver closed.");
            }
        }
    }

    // Method to handle location selection dynamically
    public static void handleLocationSelection(WebDriver driver, String location, String pincode) throws InterruptedException {
    	// Location Selection Logic (with dynamic pincode)
        WebElement locationFirstClick;
        try {
            locationFirstClick = driver.findElement(By.xpath("(//div[@class='flex w-full'])[2]"));
        } catch (Exception e) {
            locationFirstClick = driver.findElement(By.xpath("(//div[@class='flex w-full'])[1]"));
        }
        Thread.sleep(1000);
        locationFirstClick.click();

        Thread.sleep(1000);
//        WebElement sendCode;
//        try {
//            sendCode = driver.findElement(By.xpath("(//div[@class='AddressDropdown___StyledDiv-sc-i4k67t-7 eXGbTp'])[2]"));
//        } catch (Exception e) {
//            sendCode = driver.findElement(By.xpath("(//div[@class='AddressDropdown___StyledDiv-sc-i4k67t-7 eXGbTp'])[1]"));
//        }
//        Thread.sleep(1000);
//        sendCode.click();

        WebElement codeSend;
        try {
            codeSend = driver.findElement(By.xpath("(//input[@class='Input-sc-tvw4mq-0 AddressDropdown___StyledInput-sc-i4k67t-8 hpyysx eQvECn'])[2]"));
        } catch (Exception e) {
            codeSend = driver.findElement(By.xpath("(//input[@class='Input-sc-tvw4mq-0 AddressDropdown___StyledInput-sc-i4k67t-8 hpyysx eQvECn'])[1]"));
        }
        Thread.sleep(1000);
        codeSend.click();
        Thread.sleep(1000);

        // Assuming the pincode is in the third column (index 2) of the input data
      //  String pincode = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : ""; // Read the pincode from the input sheet

        if (!pincode.isEmpty()) {
            // Proceed to set the location using the pincode from the input data
            codeSend.sendKeys(pincode); // Use the pincode fetched from input data
            Thread.sleep(1000);

            // Select the location from the suggestions
            WebElement clickLocationFirst = driver.findElement(By.xpath("//ul[@class='overscroll-contain p-2.5']//li[1]"));
            clickLocationFirst.click();
            Thread.sleep(1000);
        } else {
            System.out.println("No pincode found for location: Skipping location setting.");
        }
    }

    // Scraping methods for product details
    public static List<WebElement> scrapeCategoryLinks(WebDriver driver) {
        return driver.findElements(By.xpath("//section[@class='z-10 ']//li"));
    }

    public static String getProductUrl(WebElement categoryLink) {
        return categoryLink.findElement(By.xpath(".//div[@class='DeckImage___StyledDiv-sc-1mdvxwk-1 jbskZj']//a")).getAttribute("href");
    }

    public static String getProductBrandName(WebElement categoryLink) {
        return categoryLink.findElement(By.xpath(".//h3//span[@class='Label-sc-15v1nk5-0 BrandName___StyledLabel2-sc-hssfrl-1 gJxZPQ keQNWn']")).getText();
    }

    public static String getProductName(WebElement categoryLink) {
        return categoryLink.findElement(By.xpath(".//h3[@class='block m-0 line-clamp-2 font-regular text-base leading-sm text-darkOnyx-800 pt-0.5 h-full']")).getText();
    }

    public static String getOriginalMrp(WebElement categoryLink) {
        String mrp = categoryLink.findElement(By.xpath(".//span[@class='Label-sc-15v1nk5-0 Pricing___StyledLabel2-sc-pldi2d-2 gJxZPQ hsCgvu']")).getText();
        // Remove the rupee symbol and any extra spaces or commas
        mrp = mrp.replaceAll("₹", "").replaceAll(",", "").trim();
        return mrp;
    }

    public static String getSpValue(WebElement categoryLink) {
        String sp = categoryLink.findElement(By.xpath(".//span[@class='Label-sc-15v1nk5-0 Pricing___StyledLabel-sc-pldi2d-1 gJxZPQ AypOi']")).getText();
        // Remove the rupee symbol and any extra spaces or commas
        sp = sp.replaceAll("₹", "").replaceAll(",", "").trim();
        return sp;
    }

    public static String getOriginalUom(WebElement categoryLink, String productNameResult) {
        return categoryLink.findElement(By.xpath("//h3[contains(text(),'" + productNameResult + "')]/ancestor::a/following-sibling::div[2]//span")).getText();
    }

    public static String getOffer(WebElement categoryLink, String productNameResult) {
        return categoryLink.findElement(By.xpath(".//h3[contains(text(),'" + productNameResult + "')]/ancestor::h3/parent::div/div[1]//div[contains(@class,'Offers')]//span[contains(text(),'% ')]")).getText();
    }
    
    public static String extractProductCode(String url) {
        // Regular expression to match the product code in the URL
        String regex = "product/(\\d+)";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(url);
        if (matcher.find()) {
            return matcher.group(1); // Return the matched product code
        }
        return "NA";  // If no match is found, return "NA"
    }

    // Scroll and scrape method
    public static void scrollAndScrape(WebDriver driver) throws InterruptedException {
        JavascriptExecutor js = (JavascriptExecutor) driver;

        // First scroll: Scroll down by 1000px to load more products.
        js.executeScript("window.scrollBy(0, 1000)");
        Thread.sleep(3000);

        // Second scroll: Scroll by 1000px increments to load more items gradually.
        long scrollHeight = (Long) js.executeScript("return document.body.scrollHeight");
        for (long i = 1000; i < scrollHeight; i += 1000) {
            js.executeScript("window.scrollBy(0, 1000)");
            Thread.sleep(2000); // Wait for the content to load
        }
    }
}
