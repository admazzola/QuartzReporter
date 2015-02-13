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
		
		HashSet<Row> previousTechConnectRows = previousSalesSheet.getRowsWithProductCode(ProductCodes.TECHCONNECT);
		HashSet<Row> currentTechConnectRows = currentSalesSheet.getRowsWithProductCode(ProductCodes.TECHCONNECT);
		
		HashSet<Row> resultRows = new HashSet<Row>();
		
		Row titleRow = resultsSheet.createRow( 0 );
		titleRow.createCell(ResultColumns.CustNo).setCellValue("BPID");
		titleRow.createCell(ResultColumns.CustName).setCellValue("Name");
		titleRow.createCell(ResultColumns.PrevSales).setCellValue("PrevSales");
		titleRow.createCell(ResultColumns.CurrSales).setCellValue("CurrentSales");
		
		
		int previousCustomerNoColumn = QuartzReporter.getColumnContainingString(previousSalesSheet.getTitles(), "CustomerNo");
		int previousCustomerNameColumn = QuartzReporter.getColumnContainingString(previousSalesSheet.getTitles(), "CustomerName");
		int previousCustomerSalesColumn = QuartzReporter.getColumnContainingString(previousSalesSheet.getTitles(), "Sales");
		
		

		for(Row myrow : previousTechConnectRows)
		{

			Cell custNoCell = myrow.getCell( previousCustomerNoColumn );
			Cell custNameCell = myrow.getCell( previousCustomerNameColumn );
			Cell custSalesCell = myrow.getCell( previousCustomerSalesColumn );
			
			
			
			
			System.out.println("rows: "+ resultsSheet.getLastRowNum());
			Row newRow = resultsSheet.createRow( resultsSheet.getLastRowNum() +1 );
			resultRows.add(newRow); //add reference to my list for easy access
			
			copyCell(custNoCell  ,   newRow.createCell(ResultColumns.CustNo));
			copyCell(custNameCell  ,   newRow.createCell(ResultColumns.CustName));
			copyCell(custSalesCell  ,   newRow.createCell(ResultColumns.PrevSales));
			
		}
		
		
		
		int currentCustomerNoColumn = QuartzReporter.getColumnContainingString(currentSalesSheet.getTitles(), "CustomerNo");
		int currentCustomerNameColumn = QuartzReporter.getColumnContainingString(currentSalesSheet.getTitles(), "CustomerName");
		int currentCustomerSalesColumn = QuartzReporter.getColumnContainingString(currentSalesSheet.getTitles(), "Sales");
		
		
		for(Row myrow : currentTechConnectRows)
		{
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
				
				copyCell(custNoCell  ,   newRow.createCell(ResultColumns.CustNo));
				copyCell(custNameCell  ,   newRow.createCell(ResultColumns.CustName));
				copyCell(custSalesCell  ,   newRow.createCell(ResultColumns.CurrSales));
				
			}
		//add data to 4
		}
		
		
		
		
		
		//Export the Results File
		
		String outputFilePath = QuartzReporter.getOutputPathToNewFile("Select the Output File");
		
		FileOutputStream outputFile = new FileOutputStream(new File( outputFilePath ));
		outputbook.write(outputFile);
		
	}
	
	
	
	 
	 
	private void copyRowToSheet(HSSFSheet resultsSheet, Row myrow) {
		
		System.out.println("copying row "+myrow);

			HSSFRow newRow = resultsSheet.createRow( resultsSheet.getLastRowNum()+1 );
			
			for (Iterator<Cell> cells = myrow.iterator(); cells.hasNext(); ) {
			    Cell oldCell = cells.next();
			    
			    
			    
			    HSSFCell newCell = newRow.createCell( oldCell.getColumnIndex() );	
			    
			    copyCell(oldCell , newCell);
			    
			    
			}
			
		
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
				
				int ProdCatColumn = QuartzReporter.getColumnContainingString(titles, "ProdCategory");
				
								
			}
			
			public Row getTitles() {
				return titles;
			}

			HashSet<Row> getRowsWithProductCode(String code)
			{
				
				
				
				HashSet<Row> techConnectSales = selectRowsWhere(mysheet, "PriceType", ProductCodes.TECHCONNECT   );
				
				
				
				return techConnectSales;
				
				
			}
			

			
		}
	 
	
	class ResultColumns
	{
		public static final int CustNo = 0;
		public static final int CustName = 1;
		public static final int PrevSales = 2;
		public static final int CurrSales = 3;
	}
	
}
