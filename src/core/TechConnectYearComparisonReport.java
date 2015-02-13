package core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class TechConnectYearComparisonReport extends Report{
	

	
	protected void init() throws Exception{
		
		String previousFilePath = QuartzReporter.getPathFromFileChooser("Select the Previous Report");
		
		String currentFilePath = QuartzReporter.getPathFromFileChooser("Select the Current Report");

		HSSFWorkbook outputbook = new HSSFWorkbook();//need to save this eventually..
				
		HSSFSheet resultsSheet = outputbook.createSheet("QuartzReportResults");
		 
		
		FileInputStream previousFile = new FileInputStream(new File( previousFilePath ));
		FileInputStream currentFile = new FileInputStream(new File( currentFilePath ));
		
		
				//Get the workbook instance for XLS file 
		HSSFWorkbook previousbook = new HSSFWorkbook(previousFile);
		HSSFWorkbook currentbook = new HSSFWorkbook(currentFile);
		
		
		TechConnectSalesSheet previousSalesSheet = new TechConnectSalesSheet( previousbook.getSheet("Sales Detail") );
		TechConnectSalesSheet currentSalesSheet = new TechConnectSalesSheet( currentbook.getSheet("Sales Detail") );
		
		HashSet<Row> previousTechConnectRows = previousSalesSheet.getRowsWithColumnValue("PriceType",ProductCodes.TECHCONNECT);
		HashSet<Row> currentTechConnectRows = currentSalesSheet.getRowsWithColumnValue("PseudoPPT",ProductCodes.TECHCONNECT);
		
		HashSet<Row> resultRows = new HashSet<Row>();
		
		Row titleRow = resultsSheet.createRow( 0 );
		titleRow.createCell(ResultColumns.ProdCode).setCellValue("Product Code");
		titleRow.createCell(ResultColumns.CustNo).setCellValue("BPID");
		titleRow.createCell(ResultColumns.CustName).setCellValue("Name");
		titleRow.createCell(ResultColumns.PrevSales).setCellValue("Prev Sales");
		titleRow.createCell(ResultColumns.CurrSales).setCellValue("Current Sales");
		titleRow.createCell(ResultColumns.NetSales).setCellValue("Net Sales");
		
		int previousProductCodeColumn = QuartzReporter.getColumnContainingString(previousSalesSheet.getTitles(), "ProdCategory");
		int previousCustomerNoColumn = QuartzReporter.getColumnContainingString(previousSalesSheet.getTitles(), "CustomerNo");
		int previousCustomerNameColumn = QuartzReporter.getColumnContainingString(previousSalesSheet.getTitles(), "CustomerName");
		int previousCustomerSalesColumn = QuartzReporter.getColumnContainingString(previousSalesSheet.getTitles(), "Sales");
		
		for(Row myrow : previousTechConnectRows)
		{

			Cell prodCatCell = myrow.getCell( previousProductCodeColumn );
			Cell custNoCell = myrow.getCell( previousCustomerNoColumn );
			Cell custNameCell = myrow.getCell( previousCustomerNameColumn );
			Cell custSalesCell = myrow.getCell( previousCustomerSalesColumn );
			
						
			System.out.println("rows: "+ resultsSheet.getLastRowNum());
			Row newRow = resultsSheet.createRow( resultsSheet.getLastRowNum() +1 );
			resultRows.add(newRow); //add reference to my list for easy access
			
			copyCell(prodCatCell  ,   newRow.createCell(ResultColumns.ProdCode));
			copyCell(custNoCell  ,   newRow.createCell(ResultColumns.CustNo));
			copyCell(custNameCell  ,   newRow.createCell(ResultColumns.CustName));
			copyCell(custSalesCell  ,   newRow.createCell(ResultColumns.PrevSales));
			
		}
		
		
		int currentProductCodeColumn = QuartzReporter.getColumnContainingString(currentSalesSheet.getTitles(), "ProdCategory");
		int currentCustomerNoColumn = QuartzReporter.getColumnContainingString(currentSalesSheet.getTitles(), "CustomerNo");
		int currentCustomerNameColumn = QuartzReporter.getColumnContainingString(currentSalesSheet.getTitles(), "CustomerName");
		int currentCustomerSalesColumn = QuartzReporter.getColumnContainingString(currentSalesSheet.getTitles(), "Sales");
		
		
		for(Row myrow : currentTechConnectRows)
		{
			Cell prodCatCell = myrow.getCell( currentProductCodeColumn );
			Cell custNoCell = myrow.getCell( currentCustomerNoColumn );
			Cell custNameCell = myrow.getCell( currentCustomerNameColumn );
			Cell custSalesCell = myrow.getCell( currentCustomerSalesColumn );
			
			//find row w matching customer number
			
			Row matchingRow = null;
			
			for(Row resultrow : resultRows)
			{
				if( custNoCell.getNumericCellValue() ==   resultrow.getCell(ResultColumns.CustNo).getNumericCellValue()  )
				{
					matchingRow = resultrow;
				}
			}
			//put in current sales data
			
			if(matchingRow != null)
			{
				copyCell(custSalesCell  ,   matchingRow.createCell(3));
			
			}
			else
			{
				Row newRow = resultsSheet.createRow( resultsSheet.getLastRowNum() +1);
				resultRows.add(newRow); //add reference to my list for easy access
				
				copyCell(prodCatCell  ,   newRow.createCell(ResultColumns.ProdCode));
				copyCell(custNoCell  ,   newRow.createCell(ResultColumns.CustNo));
				copyCell(custNameCell  ,   newRow.createCell(ResultColumns.CustName));
				copyCell(custSalesCell  ,   newRow.createCell(ResultColumns.CurrSales));
				
			}
		//add data to 4
		}
		
		
		//calculate other data like net sales
		for(Row myrow : resultRows)
		{
			
			
			
			int currSales = 0;
			int prevSales = 0;
			
			if(myrow.getCell(ResultColumns.CurrSales)!=null)
			currSales = (int) myrow.getCell(ResultColumns.CurrSales).getNumericCellValue();
			
			if(myrow.getCell(ResultColumns.PrevSales)!=null)
			prevSales = (int) myrow.getCell(ResultColumns.PrevSales).getNumericCellValue();
			
			int netsales = currSales - prevSales;
			
			myrow.createCell(ResultColumns.NetSales).setCellValue(netsales );
			
		}
		
		
		
		
		//Export the Results File
		
		String outputFilePath = QuartzReporter.getOutputPathToNewFile("Select the Output File");
		
		FileOutputStream outputFile = new FileOutputStream(new File( outputFilePath ));
		outputbook.write(outputFile);
		
	}
	
	
	
	 





		private void copyCell(Cell oldCell, Cell cell) {
			
			if (oldCell == null) {
                cell = null;
                return;
            }

			cell.setCellType( oldCell.getCellType() );
			
			// Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    cell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    cell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    cell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    cell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    cell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    cell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
			
		
		}





		class TechConnectSalesSheet extends SalesSheet
		{
			
			
			
			TechConnectSalesSheet(HSSFSheet sheet)
			{
				mysheet = sheet;
				

				//Get iterator to all the rows in current sheet
				titles = mysheet.getRow(0);
				mysheet.removeRow(titles);
				
				
								
			}
			
			public Row getTitles() {
				return titles;
			}

			HashSet<Row> getRowsWithColumnValue(String title, String value)
			{				
				HashSet<Row> techConnectSales = selectRowsWhere(mysheet, title,value   );
												
				return techConnectSales;
				
				
			}
			

			
		}
	 
	
	class ResultColumns
	{
		public static final int ProdCode = 0;
		public static final int CustNo = 1;
		public static final int CustName = 2;
		public static final int PrevSales = 3;
		public static final int CurrSales = 4;
		public static final int NetSales = 5;
	}
	
}
