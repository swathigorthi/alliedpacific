package alliedPacific.writer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import alliedPacific.reader.ExcelReader;
import alliedPacific.reader.InvoiceReader;

public class SalesOrderUpdater {
	private File currentSalesOrder;
	String invoiceNumber; 
	private Map<String, Double> itemNumberToQtyMap; 
	
	public SalesOrderUpdater(String currentSalesOrderPath, String pdfInvoicePath){
		this.currentSalesOrder = new File(currentSalesOrderPath);
		InvoiceReader reader = new InvoiceReader(pdfInvoicePath);
		invoiceNumber = reader.getInvoiceNumber();
		itemNumberToQtyMap = reader.getItemNoToQtyMap();
	}
	
	public void updateSalesOrder(){
		try {
			ExcelReader reader = new ExcelReader(currentSalesOrder);
		    int colNo = reader.getRowToInsertNewInvoice();
		    int newRow = reader.getRowToInsertNewItemNumber();
		    final FileInputStream  fis = new FileInputStream(currentSalesOrder);
		    final XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
		    XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			 //Write invoice number
			Row row = mySheet.getRow(0);
			Cell cell = row.getCell(colNo);
			cell.setCellValue(invoiceNumber);
			for(String itemNo : itemNumberToQtyMap.keySet()){
			    Integer rowNo = reader.getRowNumber(itemNo);
			    row = rowNo!=null? mySheet.getRow(rowNo) : mySheet.getRow(newRow);
			    if(rowNo==null){ newRow++;}
			    cell = row.getCell(colNo);
			    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			    cell.setCellValue(itemNumberToQtyMap.get(itemNo));
			}
			//reevaluate all formula cells
		    XSSFFormulaEvaluator.evaluateAllFormulaCells(myWorkBook);
			FileOutputStream os = new FileOutputStream(currentSalesOrder);
			myWorkBook.write(os);
			os.close();
			myWorkBook.close();
			fis.close();
		} 
		catch (Exception e) {
			System.out.println("Could not update sales order with item qty from invoice "+invoiceNumber);
		}
		System.out.println("Updated sales order with item qty from invoice "+ invoiceNumber);
	}
	
//	public static void main(String[] args) throws Exception {
//		File myFile = new File("C://Users/sgorthi/Documents/SalesOrder30.xlsx");
//		ExcelReader reader = new ExcelReader(myFile);
//		String invoiceNumber = "1046362";
//		Map<String, String>  itemNumberToQtyMap = new HashMap<>();
//		itemNumberToQtyMap.put("GA100-1A4", "2");
//		itemNumberToQtyMap.put("GA110MB-1ACR", "10");
////		
////		System.out.println(reader.getRowToInsertNewItemNumber());
////		System.out.println(reader.getRowToInsertNewInvoice());
////	
////		System.out.println(reader.getRowNumber("BA-110-1A"));
//		
//		final FileInputStream  fis = new FileInputStream(myFile);
//		final XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
//		 XSSFSheet mySheet = myWorkBook.getSheetAt(0);
//		 int colNo = reader.getRowToInsertNewInvoice();
//		 int newRow = reader.getRowToInsertNewItemNumber();
//		 //Write invoice number
//		Row row = mySheet.getRow(0);
//		Cell cell = row.getCell(colNo);
//		cell.setCellValue(invoiceNumber);
//		for(String itemNo : itemNumberToQtyMap.keySet()){
//		    Integer rowNo = reader.getRowNumber(itemNo);
//		    row = rowNo!=null? mySheet.getRow(rowNo) : mySheet.getRow(newRow);
//		    if(rowNo==null){ newRow++;}
//		    cell = row.getCell(colNo);
//		    cell.setCellType(Cell.CELL_TYPE_NUMERIC);
//		    cell.setCellValue(Double.valueOf(itemNumberToQtyMap.get(itemNo)));
//		}
//		//reevaluate all formula cells
//	    XSSFFormulaEvaluator.evaluateAllFormulaCells(myWorkBook);
//		
//		FileOutputStream os = new FileOutputStream(myFile);
//		myWorkBook.write(os);
//		os.close();
//		myWorkBook.close();
//		fis.close();
//	}

}
