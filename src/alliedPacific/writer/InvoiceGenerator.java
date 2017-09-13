package alliedPacific.writer;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import alliedPacific.reader.ExcelReader;

public class InvoiceGenerator {
	File currentSalesOrder ;
	String invoiceToBeCreated ; 
	Map<String, Double> currentInventoryItemQtyMap;
	Map<String, Double> currentInventoryItemPriceMap;
	
	public InvoiceGenerator(String  currentSalesOrderPath, String  invoiceToBeCreatedPath){
		this.currentSalesOrder = new File(currentSalesOrderPath);
		this.invoiceToBeCreated = invoiceToBeCreatedPath;
		ExcelReader reader = new ExcelReader(currentSalesOrder);
		currentInventoryItemQtyMap = reader.getCurrentInventoryItemQtyMap();
		currentInventoryItemPriceMap = reader.getCurrentInventoryItemPriceMap();
	}
	
	public void generateInvoice() {
		String sheetName = "Sheet1";
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName) ;
		//Set up headers
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("Item No");
		cell = row.createCell(1);
		cell.setCellValue("Price");
		cell = row.createCell(2);
		cell.setCellValue("Qty");
		cell=row.createCell(3);
		cell.setCellValue("Total price");
		//Start writing item numbers and their quantities, prices
		int r = 1;
		int totalQty = 0;
		double totalPrice = 0;
		for(String itemNumber :currentInventoryItemQtyMap.keySet()){
			Double qty = currentInventoryItemQtyMap.get(itemNumber);
			if(qty!=null && qty>0){
				totalQty +=qty;
				row = sheet.createRow(r);
				cell = row.createCell(0);
				cell.setCellValue(itemNumber);
				cell = row.createCell(1);
				Double price = currentInventoryItemPriceMap.get(itemNumber);
				double newPrice = price!=null?  1.1d*price : 0d;
				cell.setCellValue(newPrice);
				cell = row.createCell(2);
				cell.setCellValue(qty);
				cell = row.createCell(3);
				double itemPrice = newPrice * qty;
				cell.setCellValue(itemPrice);
				totalPrice+=itemPrice;
				r++;
			}
		}
		//Set up totals row
		row = sheet.createRow(r);
		cell = row.createCell(0);
		cell.setCellValue("TOTALS");
		cell = row.createCell(2);
		cell.setCellValue(totalQty);
		cell = row.createCell(3);
		cell.setCellValue(totalPrice);
		//Write ouput
		try{
			FileOutputStream fileOut = new FileOutputStream(invoiceToBeCreated);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
			wb.close();
		}
		catch(Exception e){
			System.out.print("Could generate invoice"+ e.getMessage());
		}
		System.out.println("Invoice generated successfully!!");
		
	}
	
//	
//	public static void main(String[] args) throws Exception {
//		
//		File myFile = new File("C://Users/sgorthi/Documents/SalesOrder30.xlsx");
//		ExcelReader reader = new ExcelReader(myFile);
//
//
//		String sheetName = "Sheet1";//name of sheet
//
//		XSSFWorkbook wb = new XSSFWorkbook();
//		XSSFSheet sheet = wb.createSheet(sheetName) ;
//		
//		XSSFRow row = sheet.createRow(0);
//		XSSFCell cell = row.createCell(0);
//		cell.setCellValue("Item No");
//		cell = row.createCell(1);
//		cell.setCellValue("Price");
//		cell = row.createCell(2);
//		cell.setCellValue("Qty");
//		cell=row.createCell(3);
//		cell.setCellValue("Total price");
//		int r = 1;
//		int totalQty = 0;
//		double totalPrice = 0;
//		for(String itemNumber : reader.getCurrentInventoryItemQtyMap().keySet()){
//			Double qty = reader.getCurrentInventoryItemQtyMap().get(itemNumber);
//			if(qty!=null && qty>0){
//				totalQty +=qty;
//				row = sheet.createRow(r);
//				cell = row.createCell(0);
//				cell.setCellValue(itemNumber);
//				cell = row.createCell(1);
//				Double price = reader.getCurrentInventoryItemPriceMap().get(itemNumber);
//				double newPrice = price!=null?  1.1d*price : 0d;
//				cell.setCellValue(newPrice);
//				cell = row.createCell(2);
//				cell.setCellValue(qty);
//				cell = row.createCell(3);
//				double itemPrice = newPrice * qty;
//				cell.setCellValue(itemPrice);
//				totalPrice+=itemPrice;
//				r++;
//			}
//		}
//		row = sheet.createRow(r);
//		cell = row.createCell(0);
//		cell.setCellValue("TOTALS");
//		cell = row.createCell(2);
//		cell.setCellValue(totalQty);
//		cell = row.createCell(3);
//		cell.setCellValue(totalPrice);
//
//		String excelFileName = "C://Users/sgorthi/Documents/Test.xlsx";//name of excel file
//		FileOutputStream fileOut = new FileOutputStream(excelFileName);
////		//write this workbook to an Outputstream.
//		wb.write(fileOut);
//		
//		fileOut.flush();
//		fileOut.close();
//		wb.close();
//	
//
//	}

}
