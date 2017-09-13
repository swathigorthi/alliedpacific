package alliedPacific.reader;
import java.io.File;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;

public class InvoiceReader {
  private File pdfInvoice; 
  private String invoiceNumber;
  private Map<String, Double> itemNoToQtyMap = new HashMap<>();
  
  public InvoiceReader(String pdfInvoicePath){
	  this.pdfInvoice = new File(pdfInvoicePath);
	String[] splitStrs= pdfInvoicePath.split("_");
	if(splitStrs.length>0){
		invoiceNumber = splitStrs[splitStrs.length-1];
	}
	else{
		invoiceNumber="XXX";
	}
	itemNoToQtyMap = generateItemNoToQtyMap();
	  
  }
  
  public String getInvoiceNumber(){
	  return invoiceNumber;
  }
  
  public Map<String, Double> getItemNoToQtyMap(){
	  return itemNoToQtyMap;
  }
  
  private Map<String, Double> generateItemNoToQtyMap(){
	    Map<String, Double> itemNoToQtyMap = new HashMap<>();
	    try{
	        PDDocument pd = PDDocument.load(pdfInvoice);
	        PDFTextStripper stripper = new PDFTextStripper();
	        String lineSep = stripper.getLineSeparator();
	        List<String> lines = Arrays.asList(stripper.getText(pd).split(lineSep));
	        boolean foundItemLine = false;
	        Iterator<String> iterator = lines.iterator();
	        while(iterator.hasNext()){
	        	String line = iterator.next();
//	             System.out.println(line);
	        	if(line.contains("Continued")){
	        		foundItemLine = false;
	        		continue;
	        	}
	        	if(line.contains("sushildosi@gmail.com") && !foundItemLine){
	        		foundItemLine = true;
	        		continue;
	        	}
	        	if(foundItemLine){
	        		List<String> itemLineInfo = Arrays.asList(line.split(" "));
	        		if(itemLineInfo.size()>1){
	        			String itemNo = itemLineInfo.get(0);
	        			String shippedQty = itemLineInfo.get(2);
	        			//Make sure shippedQty is an integer???
	        			itemNoToQtyMap.put(itemNo,Double.valueOf(shippedQty));
	        		}
	        		else{
	        			//Make sure its always the CR item lines that split in the middle???
	        			line+=iterator.next()+" "+iterator.next();
	        			if(line.contains("Net")){
	        				break;
	        			}
	        			itemLineInfo = Arrays.asList(line.split(" "));
	        			if(itemLineInfo.size()>1){
		        			String itemNo = itemLineInfo.get(0);
		        			String shippedQty = itemLineInfo.get(2);
		        			//Make sure shippedQty is an integer???
		        			itemNoToQtyMap.put(itemNo,Double.valueOf(shippedQty));
	        			}
	        		}
	        	}
	        }
	    }
	    catch(Exception e){
	    	System.out.println("Error: Could generate itemToQtyMap from the pdf invoice  "+ e.getMessage());
	    }
	    return itemNoToQtyMap;
  }
  

//    public static void main(String[] args) throws IOException {
//        File f= new File("C:\\Invoice_1046138.pdf");
//        PDDocument pd = PDDocument.load(f);
//        PDFTextStripper stripper = new PDFTextStripper();
//        String lineSep = stripper.getLineSeparator();
//        List<String> lines = Arrays.asList(stripper.getText(pd).split(lineSep));
//        boolean foundItemLine = false;
//        Map<String, String> itemNoToQtyMap = new HashMap<>();
//        Iterator<String> iterator = lines.iterator();
//        while(iterator.hasNext()){
//        	String line = iterator.next();
////             System.out.println(line);
//        	if(line.contains("Continued")){
//        		foundItemLine = false;
//        		continue;
//        	}
//        	if(line.contains("sushildosi@gmail.com") && !foundItemLine){
//        		foundItemLine = true;
//        		continue;
//        	}
//        	if(foundItemLine){
//        		List<String> itemLineInfo = Arrays.asList(line.split(" "));
//        		if(itemLineInfo.size()>1){
//        			String itemNo = itemLineInfo.get(0);
//        			String shippedQty = itemLineInfo.get(2);
//        			//Make sure shippedQty is an integer???
//        			itemNoToQtyMap.put(itemNo,shippedQty);
//        		}
//        		else{
//        			//Make sure its always the CR item lines that split in the middle???
//        			line+=iterator.next()+" "+iterator.next();
//        			if(line.contains("Net")){
//        				break;
//        			}
//        			itemLineInfo = Arrays.asList(line.split(" "));
//        			if(itemLineInfo.size()>1){
//	        			String itemNo = itemLineInfo.get(0);
//	        			String shippedQty = itemLineInfo.get(2);
//	        			//Make sure shippedQty is an integer???
//	        			itemNoToQtyMap.put(itemNo,shippedQty);
//        			}
//        		}
//        	}
//        }
////        for(String key :itemNoToQtyMap.keySet() ){
////        	System.out.println(itemNoToQtyMap.get(key) + "  "+key);
////        }
//    }
}


