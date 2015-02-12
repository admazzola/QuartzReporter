package core;

import java.io.File;
import java.io.FileInputStream;


import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class QuartzReporter {

	public static void main(String[] args) {
		
		if(args.length == 0)
		{
			Printer.print("You must specify an argument for an input Excel '97 file!");	
			return;
		}
		
		String inputFilePath = args[0];
		
		
		try {
			init(inputFilePath);
		} catch (Exception e) {
			
			e.printStackTrace();
		}

		

	}

	private static void init(String inputFilePath) throws Exception{

		FileInputStream file = new FileInputStream(new File( inputFilePath ));
		

		//Get the workbook instance for XLS file 
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		 
		//Get first sheet from the workbook
		HSSFSheet sheet = workbook.getSheetAt(0);
		 
		//Get iterator to all the rows in current sheet
		Iterator<Row> rowIterator = sheet.iterator();
		
		Row row = rowIterator.next();
		 
		//Get iterator to all cells of current row
		Iterator<Cell> cellIterator = row.cellIterator();


		
		
	}

}
