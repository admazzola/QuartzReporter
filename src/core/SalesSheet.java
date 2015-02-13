package core;

import java.util.HashSet;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Row;

public class SalesSheet {

	
	HSSFSheet mysheet;
	Row titles;
	
	

	HashSet<Row> selectRowsWhere(HSSFSheet sheet, String columnName, String matchingValue)
	{
		HashSet<Row> selection = new HashSet<Row>();
				
	//	HSSFWorkbook workbook = sheet.getWorkbook();
		
	//	HSSFSheet newsheet = workbook.createSheet(matchingValue);
		
		int columnIndex = getColumnContainingString( columnName  );
		
		for (Iterator<Row> rowIterator = sheet.iterator(); rowIterator.hasNext(); ) {
			Row myrow = rowIterator.next();
			
			if(myrow.getCell( columnIndex  )!=null &&  matchingValue.equals(myrow.getCell( columnIndex  ).getStringCellValue())  )
			{
				selection.add(myrow);				
			}
			
		}
		
		return selection;
	}

	 int getColumnContainingString(String string) {

		return QuartzReporter.getColumnContainingString(titles,string);
	}
 
	
}
