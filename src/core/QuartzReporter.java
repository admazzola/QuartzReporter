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
		
		
		
		
		Report myReport = new TechConnectYearComparisonReport();
				
		try {
			myReport.init();
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		

	}

	static String getPathFromFileChooser(String title) {
		
		JFileChooser chooser = new JFileChooser(DEFAULT_DIRECTORY);
		chooser.setDialogTitle(title);
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
	
	static String getOutputPathToNewFile(String title) {
		
		JFileChooser chooser = new JFileChooser(DEFAULT_DIRECTORY);
		chooser.setDialogTitle(title);
	    
	    JPanel mainPanel = new JPanel();
	    
	    int returnVal = chooser.showSaveDialog(mainPanel);
	    	    
	    if(returnVal == JFileChooser.APPROVE_OPTION) {
	       System.out.println("File saved at: " +
	            chooser.getSelectedFile().getName());
	    }
		
	    String path = chooser.getSelectedFile().getAbsolutePath();
	    
	    if( !path.endsWith(".xls") )
	    {
	    	
	    	path = path + ".xls";
	    }
	    
	    
		return path;
	}



	static int getColumnContainingString(Row titles, String string) {

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
