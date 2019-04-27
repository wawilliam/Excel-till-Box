import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
/**
 * Box_Excel
 *
 * @author Wiliam Andersson
 * @version 6 mars 2017
 * 
 * Write klassen skriver ut datan i cellerna f�r varje kund och sparar det sedan
 * i respektive mapp i molnlagringen f�r att sedan kunna anv�ndas av kunderna.
 * 
 */
public class Write {


		public static void write(int amount, String url, String mapp) throws IOException {
	
		SXSSFWorkbook workbook = new SXSSFWorkbook();
		SXSSFSheet sheet = workbook.createSheet("Appendix 1");
		sheet.createFreezePane(0,0);
		
		workbook.getSheetAt(workbook.getActiveSheetIndex()).createFreezePane(0, 1);	
		
		Row row1 = sheet.createRow(0);
		
		for(int u = 0; u <Read.colNum; u++) {
		
		
			
			Cell cell1 = row1.createCell(u);
			//�kar storleken p� f�rsta raden
			row1.setHeightInPoints(35);
			
			cell1.setCellValue(Read.top[u]);
		}
	
		  
		//Inititerar en cellstill och s�tter typsnittet till fet.
		CellStyle style = workbook.createCellStyle();
		    Font font = workbook.createFont();
		    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		    style.setFont(font);
		    
		  //H�r s�tts de �versta cellerna till fet.
		    for(int i = 0; i < Read.colNum; i++)
		    {
		        row1.getCell(i).setCellStyle(style);
		    }
		
		
		    for(int j = 0; j < Read.customersLength[amount+1]; j++) 
		    {
		
		    	Row row = sheet.createRow(j+1);
		
		    	for(int i = 0; i < Read.colNum; i++)
		    	{
		
		    		Cell cell = row.createCell(i);
		    		cell.setCellValue(Read.data[amount] [j] [i].toString());
	
		    	}
		    }	
		
		
		//fryser f�rsta raden
		workbook.getSheetAt(workbook.getActiveSheetIndex()).createFreezePane(0, 1);	
		//H�r skrivs filerna ut med sitta best�mda namn enligt standard. (Excel_fil_v1_"kundnamnet".xlsx)
		workbook.write(new FileOutputStream(url + "/" + mapp + "/" + "Excel_fil_lista_" + Read.data[amount] [0]  [0] + ".xlsx"));
	
		workbook.close();
		
		
		Read.data[amount] = null;
	
	
	
		}
	
		
		
	
	
	

		
	
	
}
	
