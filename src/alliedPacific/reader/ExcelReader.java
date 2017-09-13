package alliedPacific.reader;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	private File currentSalesOrder;	
	private Map<String, Integer> itemNumberToRowNumberMap = new HashMap<>();
	private Map<Integer, String> rowNumberToItemNumberMap = new HashMap<>();
	private int rowToInsertNewItemNumber=0;
	private int rowToInsertNewInvoice =0;
	private int totalsRow = 0;
	private int totalsColumn = 0;
	private Map<String, Double> currentInventoryItemQtyMap = new HashMap<>();
	private Map<String, Double> currentInventoryItemPriceMap = new HashMap<>();
	
	public ExcelReader(File myFile){
		this.currentSalesOrder = myFile;
		process();
	}
	private void process(){
//		File myFile = new File("C://SalesOrder30.xlsx");
		try {
			final FileInputStream  fis = new FileInputStream(currentSalesOrder);
			 final XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
			 XSSFSheet mySheet = myWorkBook.getSheetAt(0);
			 Iterator<Row> rowIterator = mySheet.iterator();
			 boolean foundFirstInvoiceColumn = false;
			 while (rowIterator.hasNext()) {
				 Row row = rowIterator.next();
				 rowToInsertNewItemNumber++;
				 Iterator<Cell> cellIterator = row.cellIterator();
				 while (cellIterator.hasNext()) {
					 Cell cell = cellIterator.next();
					 switch (cell.getCellType()) {
					      case Cell.CELL_TYPE_STRING:
					    	  String cellContent = cell.getStringCellValue();
					    	  if(cell.getColumnIndex()==0 && cellContent!=null && !cellContent.isEmpty() && !cellContent.equals("ITEM NO")){
					    		  itemNumberToRowNumberMap.put(cellContent.replaceAll("-", ""), cell.getRowIndex());
					    		  rowNumberToItemNumberMap.put(cell.getRowIndex(), cellContent.replaceAll("-", ""));
					    	  }
					    	  else if(cellContent.equals("Total")){
					    		  totalsRow = cell.getRowIndex();
					    		  totalsColumn = cell.getColumnIndex();
					    	  }
						      break;
					      case Cell.CELL_TYPE_NUMERIC:
					    	  double doubleCellContent = cell.getNumericCellValue();
					    	  if(cell.getRowIndex()==0){
					    		  if(!foundFirstInvoiceColumn){
					    			  rowToInsertNewInvoice = cell.getColumnIndex()+1;
					    			  foundFirstInvoiceColumn = true;
					    		  }
					    		  else if(rowToInsertNewInvoice == cell.getColumnIndex()){
					    			  rowToInsertNewInvoice++;
					    		  }
					    	  }
					    	  else if(cell.getColumnIndex()==1){
				    			  currentInventoryItemPriceMap.put(rowNumberToItemNumberMap.get(cell.getRowIndex()),doubleCellContent );
				    		  }
					    	 
						      break;
						  case Cell.CELL_TYPE_FORMULA:
							  if(cell.getRowIndex()>totalsRow+1 && totalsColumn!=0 && cell.getColumnIndex()== totalsColumn){
					    		  currentInventoryItemQtyMap.put(rowNumberToItemNumberMap.get(cell.getRowIndex()),cell.getNumericCellValue() );
					    	  }
							  break;
						  default :
					  }
				 }
				 myWorkBook.close();
				 fis.close();
			 }
		} 
		catch (Exception e) {
			System.out.println("COULD NOT PROCESS FILE "+ currentSalesOrder.getName());
		}
		System.out.println("Successfully processed file "+ currentSalesOrder.getName());
	}
	
	
	public File getMyFile() {
		return currentSalesOrder;
	}
	public Integer getRowNumber(String itemNumber) {
		return itemNumberToRowNumberMap.get(itemNumber.replaceAll("-", ""));
	}
	public int getRowToInsertNewItemNumber() {
		return rowToInsertNewItemNumber;
	}
	public int getRowToInsertNewInvoice() {
		return rowToInsertNewInvoice;
	}
	
	public int getTotalsColumn(){
		return totalsColumn;
	}

	public Map<String, Double> getCurrentInventoryItemQtyMap(){
		return currentInventoryItemQtyMap;
	}
	
	public Map<String, Double> getCurrentInventoryItemPriceMap(){
		return currentInventoryItemPriceMap;
	}
//	public static void main(String[] args) throws Exception {
//		File myFile = new File("C://SalesOrder30.xlsx");
//		FileInputStream fis = new FileInputStream(myFile);
//		 XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
//		 XSSFSheet mySheet = myWorkBook.getSheetAt(0);
//		 Iterator<Row> rowIterator = mySheet.iterator();
//				 while (rowIterator.hasNext()) {
//			 Row row = rowIterator.next();
//			 Iterator<Cell> cellIterator = row.cellIterator();
//			 while (cellIterator.hasNext()) {
//				 
//				 Cell cell = cellIterator.next();
//				 CellAddress address =cell.getAddress();
//				 address.
//				  switch (cell.getCellType()) {
//				  case Cell.CELL_TYPE_STRING:
//					  System.out.print(cell.getStringCellValue() + "\t");
//					  break;
//				  case Cell.CELL_TYPE_NUMERIC:
//					  System.out.print(cell.getNumericCellValue() + "\t");
//					  break;
//					  case Cell.CELL_TYPE_BOOLEAN:
//						  System.out.print(cell.getBooleanCellValue() + "\t");
//						  break;
//					  default :
//						  
//				  }
//			 }
//			 System.out.println("");
//		 }
//		
//
//		
//
//	}

}
