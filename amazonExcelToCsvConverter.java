package csv;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.BufferedWriter;
import java.io.IOException;

public class amazonExcelToCsvConverter {
    public static void main(String[] args) {
       //INPUT
    	String excelFilePath = ".\\S3 Input\\S3 Format Input data.xlsx";
    	
         //OUTPUT
    	 String csvFilePath = ".\\S3 OUTPUT\\04Dec2024_gr_nw_kvi_mapping_MyntraSkin+FaceCare.csv";
         
    	
//        04Dec2024_gr_nw_kvi_mapping_AmazonSkin+FaceCare             	04Dec2024_gr_nw_kvi_mapping_AmazonGurgaon
//        04Dec2024_gr_nw_kvi_mapping_BigBasketSkin+FaceCare			04Dec2024_gr_nw_kvi_mapping_AmazonDiapers
//        04Dec2024_gr_nw_kvi_mapping_FlipkartSkin+FaceCare				04Dec2024_gr_nw_kvi_mapping_AmazonNew
//        04Dec2024_gr_nw_kvi_mapping_MyntraSkin+FaceCare				04Dec2024_gr_nw_kvi_mapping_AmazonBedsheet
//        04Dec2024_gr_nw_kvi_mapping_NykaaSkin+FaceCare				04Dec2024_gr_nw_kvi_mapping_AmazonPharma
//        04Dec2024_gr_nw_kvi_mapping_PurplleSkin+FaceCare				04Dec2024_gr_nw_kvi_mapping_D2C
//        04Dec2024_gr_nw_kvi_mapping_SwiggySkin+FaceCare				04Dec2024_gr_nw_kvi_mapping_FirstCryDiapers
//    	   04Dec2024_gr_nw_kvi_mapping_ZeptoSkin+FaceCare				04Dec2024_gr_nw_kvi_mapping_FirstCryGurgaon
//        																04Dec2024_gr_nw_kvi_mapping_ZeptoPharma
//			04Dec2024_gr_nw_kvi_mapping_ApolloPharmacy					04Dec2024_gr_nw_kvi_mapping_AmazonGMCategories	    			
//			04Dec2024_gr_nw_kvi_mapping_SwiggyPharma
//			04Dec2024_gr_nw_kvi_mapping_Tata_1mgPharma	
//    	 	
  
    	 
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis);
             BufferedWriter writer = new BufferedWriter(new FileWriter(csvFilePath))) {

            Sheet sheet = workbook.getSheetAt(0); // Get the first sheet

            // Write customized headers to CSV
            
            // writer.write("Pid,City,Title,Size,Multiplier,Amazon_ProductCode,Amazon_MRP,Amazon_Selling_Price,Amazon_Availability,Amazon_UOM"); //Amazon
            
          //    writer.write("Pid,City,Title,Size,Multiplier,FirstCry_ProductCode,FirstCry_MRP,FirstCry_Selling_Price,FirstCry_Availability,FirstCry_UOM"); //Firstcry
            
            writer.write("Pid,City,Title,Size,Multiplier,Myntra_ProductCode,Myntra_MRP,Myntra_Selling_Price,Myntra_Availability,Myntra_UOM"); //Myntra
            
          //   writer.write("Pid,City,Title,Size,Multiplier,Nykaa_ProductCode,Nykaa_MRP,Nykaa_Selling_Price,Nykaa_Availability,Nykaa_UOM"); //Nykaa
            
         //   writer.write("Pid,City,Title,Size,Multiplier,Swiggy_ProductCode,Swiggy_MRP,Swiggy_Selling_Price,Swiggy_Availability,Swiggy_UOM"); //Swiggy
            
            //  writer.write("Pid,City,Title,Size,Multiplier,Zepto_ProductCode,Zepto_MRP,Zepto_Selling_Price,Zepto_Availability,Zepto_UOM"); //Zepto
            
           // writer.write("Pid,City,Title,Size,Multiplier,Purplle_ProductCode,Purplle_MRP,Purplle_Selling_Price,Purplle_Availability,Purplle_UOM"); //Purplle
            
          //  writer.write("Pid,City,Title,Size,Multiplier,BigBasket_ProductCode,BigBasket_MRP,BigBasket_Selling_Price,BigBasket_Availability,BigBasket_UOM"); //BigBasket
            
         //    writer.write("Pid,City,Title,Size,Multiplier,Flipkart_ProductCode,Flipkart_MRP,Flipkart_Selling_Price,Flipkart_Availability,Flipkart_UOM"); //Flipkart
            
            // writer.write("Pid,City,Title,Size,Multiplier,ApolloPharmacy_ProductCode,ApolloPharmacy_MRP,ApolloPharmacy_Selling_Price,ApolloPharmacy_Availability,ApolloPharmacy_UOM"); //Apollo
            
            // writer.write("Pid,City,Title,Size,Multiplier,Tata_1mg_ProductCode,Tata_1mg_MRP,Tata_1mg_Selling_Price,Tata_1mg_Availability,Tata_1mg_UOM");//Tata1mg
           
            //    writer.write("Pid,City,Title,Size,Multiplier,D2C_ProductCode,D2C_MRP,D2C_Selling_Price,D2C_Availability,D2C_UOM");//D2C
            
            writer.newLine();

            // Write data rows to CSV
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) { // Start from the second row
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    StringBuilder rowData = new StringBuilder();

                    // Define cells
                    Cell inputIdCell = row.getCell(0);
                    Cell cityCell = row.getCell(1);
                    Cell titleCell = row.getCell(2);
                    Cell sizeCell = row.getCell(3);
                    Cell multiplierCell = row.getCell(10);
                    Cell productCodeCell = row.getCell(4);
                    Cell mrpCell = row.getCell(7);
                    Cell spCell = row.getCell(8);
                    Cell availabilityCell = row.getCell(11);
                    Cell uomCell = row.getCell(9);

                    // Extract values with null checks and handle special characters
                    String inputId = safeString(inputIdCell);
                    String city = safeString(cityCell);
                    String title = safeString(titleCell);
                    String size = safeString(sizeCell);
                    String multiplier = safeString(multiplierCell);
                    String productCode = safeString(productCodeCell);
                    String mrp = safeString(mrpCell);
                    String sp = safeString(spCell);
                    String availability = safeString(availabilityCell);
                    String uom = safeString(uomCell);

                    // Append values to rowData
                    rowData.append(inputId).append(",")
                           .append(city).append(",")
                           .append(title).append(",")
                           .append(size).append(",")
                           .append(multiplier).append(",")
                           .append(productCode).append(",")
                           .append(mrp).append(",")
                           .append(sp).append(",")
                           .append(availability).append(",")
                           .append(uom);

                    // Write the row data to the CSV file
                    writer.write(rowData.toString());
                    writer.newLine();
                }
            }

            System.out.println("Excel file converted to CSV successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to safely convert cell values to string
    private static String safeString(Cell cell) {
        if (cell == null) {
            return ""; // Return empty string for null cells
        }
        String value = cell.toString().trim(); // Trim whitespace
        // Enclose in quotes if it contains a comma or newline
        if (value.contains(",") || value.contains("\n")) {
            return "\"" + value + "\"";
        }
        return value;
    }
}
