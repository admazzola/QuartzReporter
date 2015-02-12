package core;

import java.io.File;
import java.io.FileInputStream;


import java.util.Iterator;

import javax.swing.JFileChooser;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class QuartzReporter {

	final static String DEFAULT_DIRECTORY = "\\\\macnet-ad\\ns-mi-usr\\mi-usr\\mazzolaa\\Quarterly Reviews";
	
	public static void main(String[] args) {
		
		/*if(args.length == 0)
		{
			Printer.print("You must specify an argument for an input Excel '97 file!");	
			return;
		}*/
		
		String inputFilePath = getPathFromFileChooser();
		
		
		try {
			init(inputFilePath);
		} catch (Exception e) {
			
			e.printStackTrace();
		}

		

	}

	private static String getPathFromFileChooser() {
		
		JFileChooser chooser = new JFileChooser(DEFAULT_DIRECTORY);
	    FileNameExtensionFilter filter = new FileNameExtensionFilter(
	        "Excel Files", "xls");
	    chooser.setFileFilter(filter);
	    
	    JPanel mainPanel = new JPanel();
	    
	    int returnVal = chooser.showOpenDialog(mainPanel);
	    if(returnVal == JFileChooser.APPROVE_OPTION) {
	       System.out.println("You chose to open this file: " +
	            chooser.getSelectedFile().getName());
	    }
		
		return chooser.getSelectedFile().getAbsolutePath();
	}

	private static void init(String inputFilePath) throws Exception{

		FileInputStream file = new FileInputStream(new File( inputFilePath ));
		
				//Get the workbook instance for XLS file 
		HSSFWorkbook inputbook = new HSSFWorkbook(file);
		
		
		
		
		HSSFWorkbook outputbook = new HSSFWorkbook();//need to save this eventually..
				
		HSSFSheet quartzSheet = outputbook.createSheet("QuartzReport");
		 


		
		
		//Get first sheet from the workbook
		HSSFSheet salessheet = inputbook.getSheet("Sales Detail");
		
		
		//Get iterator to all the rows in current sheet
		Row titles = salessheet.getRow(0);		
		salessheet.removeRow(titles);
		
		int ProdCatColumn = getColumnContainingString(titles, "ProdCategory");
		
		for (Iterator<Row> rowIterator = salessheet.iterator(); rowIterator.hasNext(); ) {
			Row myrow = rowIterator.next();
		    // 1 - can call methods of element
		    // 2 - can use iter.remove() to remove the current element from the list

		    // ...
			
					
			Cell prodCategoryCell = myrow.getCell( ProdCatColumn );
			System.out.println( prodCategoryCell.getStringCellValue()  );
			
		}
		
	
		
	//	System.out.println("val" +   cell.getStringCellValue() );
		
		
	}

	private static int getColumnContainingString(Row titles, String string) {

		for (Iterator<Cell> cellIterator = titles.iterator(); cellIterator.hasNext(); ) {
			Cell title = cellIterator.next();
			if( string.equals(title.getStringCellValue()) )
			{
				return title.getColumnIndex();
			}
			
		}
		
		System.err.println("No columns found named "+ string);
		return 0;
	}

}
