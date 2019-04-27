import java.io.File;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
/**
 * Box_Excel
 *
 * @author Wiliam Andersson
 * @version 6 mars 2017
 * 
 * Denna klassen hanterar läsningen av Excel-filen. 
 * Den tar ut datan ifrån varje "kund" genom att kolla första kolummen efter förändringar.
 * Då en förändring intäffar, sparas datan för kunden.
 * 
 */
public class Read {
	
	
	
	//Antal kunder
	int customerAmount = 0;
	//Tillfälligt antal rader per kund
	int customerCellCounter = 0;
	//Index variabel 
	int i = 0;	
	//längderna på kunderna (MAX 100 kunder)
	static int customersLength[] = new int[100];
	//Används för att kolla att raderna på första kolummen är lika.
	boolean customerEqual = true;
	//Stoppar iteratioen
	boolean fortsatt = true;
	//föregående cell
	String cell1 = "";
	//eftergående cell
	String cell2 = "";
	//massdata dim 1: kunden dim 2: kolumner dim 3: rader
	static String[] [] [] data;
	//översta raden; titel, namn, kundnr, etc.
	static String [] top;
	//antal rader
	int rowNum;
	//antal kolummner
	static int colNum;
	//url:n för Excel-filen
	static String input = "";

	
		

		Read() throws IOException, EncryptedDocumentException, InvalidFormatException 
		{
	
		//start värde
		customersLength[0] = 0;
		read();
		}
		
		public void read() throws IOException, EncryptedDocumentException, InvalidFormatException 
		{
		
			//Excelboken initieras
			Workbook workbook = WorkbookFactory.create(new File(input));
			
			
			//Arbetsblad initieras
			Sheet sheet = workbook.getSheetAt(0);
			
			
			rowNum = sheet.getLastRowNum() + 1;
			colNum = sheet.getRow(0).getLastCellNum();
			data = new String[60] [5000] [30];
			top = new String [100];
			
			
	
			//Sparar översta raden, eftersom den ska finnas i alla excel-filer.
			Row rowTop = sheet.getRow(0);
			for(int y = 0; y < colNum ; y++ ) 
			{
				Cell cellTop = rowTop.getCell(y);
				
				top[y] = cellTop.toString();
			}
			
			
			
			//Här börjar iterationen, 
			while(fortsatt) 
			{ 
			
				i = i + 1; //itererar rader
				
				customerEqual = true;
				customerCellCounter = 0;
				
					while (customerEqual) //hanterar längd på kund
					{
					
						//rad
						Row row = sheet.getRow(i+1);
						//kolummn 1
						Cell cell = row.getCell(0, Row.CREATE_NULL_AS_BLANK );
						
						if(cell == null || cell.toString() == "") {
							customerEqual = false;
							fortsatt = false;
						}
					
						//Cell 1 tolkas till en sträng
						cell1 = cell.toString();
						
						
						if(i == rowNum-3) {
							customerEqual = false;
							fortsatt = false;
						}
					
						//eftergående rad
						Row rowSecond = sheet.getRow(i+2);
						
						//sätt till null, ifall tom.
						Cell c = rowSecond.getCell(0, Row.CREATE_NULL_AS_BLANK);
					
						
							if(c == null) {
							customerEqual = false;
							fortsatt = false;
						
							}
					
							else{
			   	  
								Cell cellSecond = rowSecond.getCell(0);
								
								
								cell2 = cellSecond.toString();
								//Cell 2
					
								//Cells for each customer
								customerCellCounter++;
								
								//När kunden inte har fler rader
								if(!cell1.equals(cell2) ) {
									customerEqual = false;
			
								}
							}
				
					 i++;
				
		
				}
			
			//Lägger till en rad för att kunna få plats med titel, namn, datum etc. på första raden.
			customersLength[customerAmount+1] = customerCellCounter+1;
			
			//Lägger till kund, eftersom startvärdet är noll och för att underlätta slutvärdet på for-lopparna nedan.
			customerAmount++;
			

			}
			//while-loopen stannar en rad innan och därför läggs en till rad på sista kunden.
			customersLength[customerAmount] = customersLength[customerAmount] +1;
			
			
			int radNummer = 0;
			
			//Här plockas alla kunder upp läggs i ett fält 
			//med hjälp av iteratioen ovan (antal rader per kund).
			//for-loopen plockar data längds rader.
			
			//kundens data bestående av rader och kolummner
			for(int q = 0; q < customerAmount; q++)
			{
				
				//raderna
				for(int j = 0; j < customersLength[q+1]; j++ )
				{
					radNummer++;
					
					Row row = sheet.getRow(radNummer); 
					
						//kolummnerna
					for(int k = 0; k < colNum ; k++) 
					{
							
							
						//initierar cell. Ifall cell tom, sätt null.
						Cell cell = row.getCell(k, row.RETURN_NULL_AND_BLANK);
					
						//ifall cellen är null sätt "" (Vilket betyder inga tecken)
						if(cell == null) 
						{
							data [q] [j] [k] = "";
						}
						
						//Annars tolka cellen och gör det till en String
						else
						{
							cell.setCellType(Cell.CELL_TYPE_STRING);
							data [q] [j] [k] = cell.toString();
						}
							
					}	
				}
			
			}
			
			workbook.close();
		
		}
}
			




