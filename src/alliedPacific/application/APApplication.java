package alliedPacific.application;

import alliedPacific.writer.InvoiceGenerator;
import alliedPacific.writer.SalesOrderUpdater;

public class APApplication {

	public static void main(String[] args) {
		String currentSalesOrderPath = "C://Users/sgorthi/Documents/SalesOrder30.xlsx";
		//update sales Order
		String pdfInvoicePath = "C://Users/sgorthi/Documents/nvoice_1046138.pdf";
		SalesOrderUpdater updater  = new SalesOrderUpdater(currentSalesOrderPath, pdfInvoicePath);
		updater.updateSalesOrder();
		
		//Generate invoice
		String invoiceToBeCreatedPath = "C://Users/sgorthi/Documents/Invoice_0030.xlsx";
		InvoiceGenerator generator = new InvoiceGenerator(currentSalesOrderPath, invoiceToBeCreatedPath);
		generator.generateInvoice();
	}

}
