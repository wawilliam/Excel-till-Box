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
 * Denna klassen hanterar l�sningen av Excel-filen. 
 * Den tar ut datan ifr�n varje "kund" genom att kolla f�rsta kolummen efter f�r�ndringar.
 * D� en f�r�ndring int�ffar, sparas datan f�r kunden.
 * 
 */
public class Read {
	
	
	
	//Antal kunder
	int customerAmount = 0;
	//Tillf�lligt antal rader per kund
	int customerCellCounter = 0;
	//Index variabel 
	int i = 0;	
	//l�ngderna p� kunderna (MAX 100 kunder)
	static int customersLength[] = new int[100];
	//Anv�nds f�r att kolla att raderna p� f�rsta kolummen �r lika.
	boolean customerEqual = true;
	//Stoppar iteratioen
	boolean fortsatt = true;
	//f�reg�ende cell
	String cell1 = "";
	//efterg�ende cell
	String cell2 = "";
	//massdata dim 1: kunden dim 2: kolumner dim 3: rader
	static String[] [] [] data;
	//�versta raden; titel, namn, kundnr, etc.
	static String [] top;
	//antal rader
	int rowNum;
	//antal kolummner
	static int colNum;
	//url:n f�r Excel-filen
	static String input = "";

	
		

		Read() throws IOException, EncryptedDocumentException, InvalidFormatException 
		{
	
		//start v�rde
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
			
			
	
			//Sparar �versta raden, eftersom den ska finnas i alla excel-filer.
			Row rowTop = sheet.getRow(0);
			for(int y = 0; y < colNum ; y++ ) 
			{
				Cell cellTop = rowTop.getCell(y);
				
				top[y] = cellTop.toString();
			}
			
			
			
			//H�r b�rjar iterationen, 
			while(fortsatt) 
			{ 
			
				i = i + 1; //itererar rader
				
				customerEqual = true;
				customerCellCounter = 0;
				
					while (customerEqual) //hanterar l�ngd p� kund
					{
					
						//rad
						Row row = sheet.getRow(i+1);
						//kolummn 1
						Cell cell = row.getCell(0, Row.CREATE_NULL_AS_BLANK );
						
						if(cell == null || cell.toString() == "") {
							customerEqual = false;
							fortsatt = false;
						}
					
						//Cell 1 tolkas till en str�ng
						cell1 = cell.toString();
						
						
						if(i == rowNum-3) {
							customerEqual = false;
							fortsatt = false;
						}
					
						//efterg�ende rad
						Row rowSecond = sheet.getRow(i+2);
						
						//s�tt till null, ifall tom.
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
								
								//N�r kunden inte har fler rader
								if(!cell1.equals(cell2) ) {
									customerEqual = false;
			
								}
							}
				
					 i++;
				
		
				}
			
			//L�gger till en rad f�r att kunna f� plats med titel, namn, datum etc. p� f�rsta raden.
			customersLength[customerAmount+1] = customerCellCounter+1;
			
			//L�gger till kund, eftersom startv�rdet �r noll och f�r att underl�tta slutv�rdet p� for-lopparna nedan.
			customerAmount++;
			

			}
			//while-loopen stannar en rad innan och d�rf�r l�ggs en till rad p� sista kunden.
			customersLength[customerAmount] = customersLength[customerAmount] +1;
			
			
			int radNummer = 0;
			
			//H�r plockas alla kunder upp l�ggs i ett f�lt 
			//med hj�lp av iteratioen ovan (antal rader per kund).
			//for-loopen plockar data l�ngds rader.
			
			//kundens data best�ende av rader och kolummner
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
							
							
						//initierar cell. Ifall cell tom, s�tt null.
						Cell cell = row.getCell(k, row.RETURN_NULL_AND_BLANK);
					
						//ifall cellen �r null s�tt "" (Vilket betyder inga tecken)
						if(cell == null) 
						{
							data [q] [j] [k] = "";
						}
						
						//Annars tolka cellen och g�r det till en String
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
			




