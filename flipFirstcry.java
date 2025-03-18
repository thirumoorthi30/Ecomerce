package Shopping;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.remote.RemoteWebDriver;

import CommonUtility.ProductDetailsScrapping2;
import CommonUtility.UserUtility;

public class flipFirstcry {

	public static void main(String[] args) {

		RemoteWebDriver driver = null;
		int rowNum = 1;

		int ExceptionCount = 0;

		// Input data (2) -- flipkart
		// Input data -- Amazon
		
		UserUtility userUtility = new UserUtility(UserUtility.InputDataFilePath10, "Input data (2)", "Locators");
		while (rowNum < userUtility.getNoOfRows(userUtility.InputDataSheet)) {
			try {
				ProductDetailsScrapping2 pds = new ProductDetailsScrapping2();
				driver = userUtility.launchBrowser(driver, false, false, false);
				for (int i = 1; i < userUtility.getNoOfRows(userUtility.InputDataSheet); i++) {
					Row row = userUtility.InputDataSheet.getRow(rowNum);

					String InputBrandName = row.getCell(0).getStringCellValue().trim();
					String InputPid = row.getCell(1).getStringCellValue().trim();
					String InputCity = row.getCell(2).getStringCellValue().trim();
					String InputTitle = row.getCell(3).getStringCellValue().trim();
					String InputSize = row.getCell(4).getStringCellValue().trim();
					String OldProductCode = row.getCell(5).getStringCellValue().trim();
					String OldURL = row.getCell(6).getStringCellValue().trim();
					String OldName = row.getCell(7).getStringCellValue().trim();
					String OldMRP = row.getCell(8).getStringCellValue().trim();
					String OldSP = row.getCell(9).getStringCellValue().trim();
					String OldUOM = row.getCell(10).getStringCellValue().trim();
					String OldMultiplier = row.getCell(11).getStringCellValue().trim();
					String OldAvailability = row.getCell(12).getStringCellValue().trim();
					String OldOffer = row.getCell(13).getStringCellValue().trim();
					
					if (rowNum == 1) {
						List<String> header = userUtility.header("");
						userUtility.writeIntoSheet(userUtility.OutputSheet, header);
					}

					if (InputBrandName.contains("Flipkart")) {
						
						System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
						
						List<String> productDetails = pds.productDetails(driver,
								userUtility.locatorsMap.get(InputBrandName), "NA", false, OldURL, InputPid, InputBrandName);
						
						
						//driver.findElement(By.id("pincodeInputId")).clear();
						
						driver.findElement(By.id("pincodeInputId")).sendKeys(Keys.ENTER);
						
						driver.findElement(By.id("pincodeInputId")).sendKeys("500001");
						
						driver.findElement(By.className("_2P_LDn")).click();
						
						
						
						String NewProductCode = productDetails.get(0).trim();
						String NewURL = productDetails.get(1).trim();
						String NewName = productDetails.get(2).trim();
						String NewMRP = productDetails.get(3).trim();
						String NewSP = productDetails.get(4).trim();
						String NewUOM = productDetails.get(5).trim();
						String NewMultiplier = productDetails.get(6).trim();
						String NewAvailability = productDetails.get(7).trim();
						String NewOffer = productDetails.get(8).trim();
						String NewManualInterventionFlag = productDetails.get(9).trim();
						String LogFile = productDetails.get(10).trim();
						String MRPPriceValidation = "NA";
						String sellingPriceValidation = "NA";
						String ProductCodeValidation = "NA";
						String OldNameCheck = "NA";
						String CrtOffer= "NA";
						if (!OldProductCode.equals("NA")) {

							if (NewURL.contains(OldProductCode)) {
								NewProductCode = OldProductCode;
								ProductCodeValidation = "OK";
							} else {
								LogFile = LogFile + "Mismatch in the product code or URL /---------";
							}

							if (OldName.equals(NewName)) {
								if (OldMRP.equals(NewMRP)) {
									NewUOM = OldUOM;
									NewMultiplier = OldMultiplier;
									NewAvailability = OldAvailability;
								} else {
									NewUOM = OldUOM;
									NewMultiplier = OldMultiplier;
									NewAvailability = OldAvailability;
									NewManualInterventionFlag = "YES";
									LogFile = LogFile + "Old MRP and New MRP is not matching /---------";

								}
							}

							else {
								NewUOM = OldUOM;
								NewMultiplier = OldMultiplier;
								NewAvailability = OldAvailability;
								NewManualInterventionFlag = "YES";
								LogFile = LogFile + "Old Name and New Name is not matching /---------";
							}
						/*	if(OldUOM.equals(NewUOM)){
							NewUOM = OldUOM;
							NewMultiplier = OldMultiplier;
							NewAvailability = OldAvailability;
							}
							else {
								NewUOM = OldUOM;
								NewMultiplier = OldMultiplier;
								NewAvailability = OldAvailability;
							}*/
							MRPPriceValidation = OldMRP;
							sellingPriceValidation = OldSP;
							OldNameCheck = OldName;
						}
						
						
						  if(NewMRP == NewSP) { 
							  NewOffer = "NA"; 
							  }
						  else if(OldOffer.equals(NewOffer)) {
						  NewOffer = OldOffer; 
						  }
						  
						  if(NewOffer != " " && InputBrandName.contains("Amazon")) {
					  			CrtOffer = NewOffer.replace("-","").replace("%","% Off");
						  }
							  else if(NewOffer != " ") { 
								  CrtOffer = NewOffer.replaceAll("off","Off"); 
								  }
							 
						 		
						List<String> finalList = new ArrayList<String>();
						finalList.add(InputPid);
						finalList.add(InputBrandName);
						finalList.add(InputTitle);
						finalList.add(InputSize);
						finalList.add(NewProductCode);
						finalList.add(NewURL);
						finalList.add(NewName);
						finalList.add(NewMRP);
						finalList.add(NewSP);
						finalList.add(NewUOM);
						finalList.add(NewMultiplier);
						finalList.add(NewAvailability);
						finalList.add(CrtOffer);
						finalList.add(NewManualInterventionFlag);
						finalList.add(LogFile);
						finalList.add(MRPPriceValidation);
						finalList.add(sellingPriceValidation);
						finalList.add(ProductCodeValidation);
						finalList.add(OldNameCheck);
						System.out.println(finalList);
						
						

						userUtility.writeIntoSheet(userUtility.OutputSheet, finalList);
					}
					rowNum++;
				}
				userUtility.writeIntoExcel(userUtility.OutputWorkBook, userUtility.OutputDataFilePath);
				System.out.println("Scrapping is completed");
			}

			catch (Exception e) {
				System.out.println(e.getMessage());
				userUtility.writeIntoExcel(userUtility.OutputWorkBook, userUtility.OutputDataFilePath);
				try {
					driver.quit();
				} catch (Exception e1) {
					e1.printStackTrace();
				}

				ExceptionCount++;

				if (ExceptionCount > 5) {
					break;
				}
			}
		}

		try {
			driver.quit();
		} catch (Exception e1) {
			e1.printStackTrace();
		}

//		System.out.println(userUtility.locatorsMap.get("Myntra"));
//		System.out.println(userUtility.getExcelDataMap(userUtility.locatorsMap, "Myntra"));
//		System.out.println(userUtility.getExcelData(userUtility.locatorsMap, "Myntra", "Product Name"));

	}

}
