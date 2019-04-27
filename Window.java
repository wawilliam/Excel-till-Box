import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.SystemColor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.SpringLayout;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
/**
 * Box_Excel
 *
 * @author Wiliam Andersson
 * @version 6 mars 2017
 * 
 * Bakgrund/Syfte: Detta projekt blev utvecklat f�r att spara tid. Tidigare har detta gjort manuellt d� man tagit
 * en kunds lista f�rsig och sedan lagt upp det p� en molnlagring. Detta tog d� otroligt l�ngt tid och jag best�mde mig
 * d� f�r att skapa ett program som l�ste detta snabbre. Tiden f�r att g�ra detta manuellt tar ca 7 timmar, medan detta program
 * g�r det under ca 20 sekunder. 
 * (Det �r ungef�r 50 kunder samt ungef�r 5000 rader = ~ 150 000 rader. 
 * 150 000 rader * 25 kolummner = 3.75 miljoner celler) (I exempel videon �r det endast 3 kunder ca 50 rader (Se demo video)
 * 
 */
public class Window {

		private JFrame frmExcelTillMappar;
		//kollar ifall box har f�tt n�got l�nk
		private boolean boxBoolean = false;
		//kollar ifall excel har f�tt n�got l�nk
		boolean excelBoolean = false;
		//s�tter ig�ng knappen f�r att k�ra programmet
		boolean buttonOn = false;
		//r�l�nk ifr�n Filv�ljaren
		String inputSourceExcel = "";
		//r�l�nk ifr�n Filv�ljaren
		String inputSourceBox = "";
		
		
		String boxDir = "";
		String excelDir = "";
		//knapp f�r att k�ra programmet.
		JButton run = new JButton("K\u00F6r");
		private JTextField excel_field;
		private SpringLayout springLayout;
		private JButton excel_dir;
		private JTextField clo_field;
		private JButton clo_dir;
		String output = "";
		//anv�nds f�r att s�tta skrivbordet till start
		String userDir = System.getProperty("user.home");
		private JButton open = new JButton();
		private JButton cancel = new JButton();
		//Ett filter f�r Excel-filer
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Filer", "xlsx", "xls" );
		private JLabel klar;
		String directories[];
		String ordning [] = new String[100];
	
	
	
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Window window = new Window();
					window.frmExcelTillMappar.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	
	public Window() {
		initialize();
		run.setEnabled(false);
		
		clo_field = new JTextField();
		springLayout.putConstraint(SpringLayout.EAST, clo_dir, -8, SpringLayout.WEST, clo_field);
		springLayout.putConstraint(SpringLayout.WEST, clo_field, 235, SpringLayout.WEST, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.EAST, clo_field, -10, SpringLayout.EAST, frmExcelTillMappar.getContentPane());
		clo_field.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		clo_field.setEditable(false);
		frmExcelTillMappar.getContentPane().add(clo_field);
		clo_field.setColumns(10);
		
		excel_field = new JTextField();
		springLayout.putConstraint(SpringLayout.SOUTH, excel_field, -112, SpringLayout.NORTH, klar);
		springLayout.putConstraint(SpringLayout.SOUTH, clo_field, -19, SpringLayout.NORTH, excel_field);
		springLayout.putConstraint(SpringLayout.WEST, excel_field, 235, SpringLayout.WEST, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.EAST, excel_field, -10, SpringLayout.EAST, frmExcelTillMappar.getContentPane());
		excel_field.setEditable(false);
		frmExcelTillMappar.getContentPane().add(excel_field);
		excel_field.setColumns(10);
	}

	
	private void initialize() {
		frmExcelTillMappar = new JFrame();
		frmExcelTillMappar.setTitle("EXCEL TILL MAPPAR");
		frmExcelTillMappar.getContentPane().setBackground(SystemColor.menu);
		frmExcelTillMappar.setResizable(false);
		frmExcelTillMappar.setBounds(100, 100, 499, 400);
		frmExcelTillMappar.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		springLayout = new SpringLayout();
		springLayout.putConstraint(SpringLayout.WEST, run, 173, SpringLayout.WEST, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.EAST, run, -187, SpringLayout.EAST, frmExcelTillMappar.getContentPane());
		frmExcelTillMappar.getContentPane().setLayout(springLayout);
		frmExcelTillMappar.setLocationRelativeTo(null);
		
		klar = new JLabel("D\u00E5 var det klart! Alla filerna \u00E4r sparade i mapparna");
		springLayout.putConstraint(SpringLayout.SOUTH, run, -16, SpringLayout.NORTH, klar);
		springLayout.putConstraint(SpringLayout.WEST, klar, 73, SpringLayout.WEST, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.EAST, klar, -82, SpringLayout.EAST, frmExcelTillMappar.getContentPane());
		klar.setHorizontalAlignment(SwingConstants.CENTER);
		springLayout.putConstraint(SpringLayout.SOUTH, klar, -10, SpringLayout.SOUTH, frmExcelTillMappar.getContentPane());
		klar.setVisible(false);
		klar.setFont(new Font("Tahoma", Font.PLAIN, 15));
		frmExcelTillMappar.getContentPane().add(klar);
		
		JLabel Text = new JLabel("SORTERING EXCEL");
		springLayout.putConstraint(SpringLayout.NORTH, Text, 51, SpringLayout.NORTH, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.WEST, Text, 0, SpringLayout.WEST, klar);
		Text.setForeground(Color.BLACK);
		Text.setFont(new Font("Arial Black", Font.BOLD, 30));
		frmExcelTillMappar.getContentPane().add(Text);
		
		clo_dir = new JButton("Bl\u00E4ddra");
		springLayout.putConstraint(SpringLayout.SOUTH, clo_dir, -178, SpringLayout.SOUTH, frmExcelTillMappar.getContentPane());
		clo_dir.addKeyListener(new KeyAdapter() {
			
		});
		clo_dir.addActionListener(new ActionListener() {
			public void actionPerformed (ActionEvent e) {
			    
				
				
				
				//H�r skapas f�nstret till filv�ljaren f�r Box
		        JFileChooser fileChooser = new JFileChooser(userDir + "/Desktop");
				fileChooser.setDialogTitle("V�lj Huvudmappen");
				fileChooser.setApproveButtonText("V�lj Mapp");
				
				
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				if(fileChooser.showOpenDialog(open) == JFileChooser.APPROVE_OPTION) {
					
					boxDir = (fileChooser.getSelectedFile().getAbsolutePath());
					
					
					
					String textUrl = "";
					String url = "";
					
					//Omvandlar "\\" till "/" f�r att den ska kunna anv�ndas som en input
					textUrl = boxDir.replace("\\" , "/");
					//Detsamma h�r, f�r att den ska fungerara som en input
					url = textUrl.replaceFirst("C", "c");
						System.out.println(url);
						
						//Kollar ifall Box har en fils�kv�g
						if(url != "") {
							boxBoolean = true;
						}
						//denna url g�r till programmet och vidare till Write-klasssen
						inputSourceBox = url;
						//Skriv ut texten p� sk�rmen
						clo_field.setText(textUrl);
						
						
				}
				else if(fileChooser.showOpenDialog(cancel) == JFileChooser.CANCEL_OPTION){
						fileChooser.setVisible(false);
					
				}
				
				
			}
		});
		frmExcelTillMappar.getContentPane().add(clo_dir);
		
		excel_dir = new JButton("Bl\u00E4ddra");
		springLayout.putConstraint(SpringLayout.NORTH, excel_dir, 14, SpringLayout.SOUTH, clo_dir);
		springLayout.putConstraint(SpringLayout.EAST, excel_dir, 0, SpringLayout.EAST, clo_dir);
		excel_dir.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {
			    
				//Excel-filv�ljaren
				JFileChooser fc = new JFileChooser(userDir + "/Desktop");
				fc.setDialogTitle("V�lj EXCEL filen");
				fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				fc.addChoosableFileFilter(filter);
				
				//ifall knappen klickas ta fils�kv�gen f�r Excel-filen
				if(fc.showOpenDialog(open) == JFileChooser.APPROVE_OPTION) {
					
					excelDir = (fc.getSelectedFile().getAbsolutePath());
			}				
				//Anv�nds ifall anv�ndaren v�ljer att avbryta
				else {
						
				}
				
				
				String displayUrl = "";	
				String url = "";
				displayUrl = excelDir.replace("\\" , "/");
				
				url = displayUrl.replaceFirst("C", "c");
				
				//ifall b�de Box samt excel har en fils�kv�g kommer k�r knappen att aktiveras.
				if(url != "" && boxBoolean ) {
					run.setEnabled(true);
				}
				System.out.println(url);
				//Visas p� sk�rmen
				excel_field.setText(displayUrl);
				
				//Denna g�r till Read-klassen f�r att hitta Excel-filen.
				Read.input = url;
		
			}
			
		});
		frmExcelTillMappar.getContentPane().add(excel_dir);
		run.setFont(new Font("Tahoma", Font.PLAIN, 16));
		run.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				
				try {
					//Anrop
					Read readClass = new Read();
					
			
					
					//Sorterar upp kundmapparna i den ordningen som dem ligger i excel-filen.
					File file = new File(inputSourceBox);
					directories = file.list(new FilenameFilter() {
				  
						public boolean accept(File current, String name) {
							return new File(current, name).isDirectory();
				    
					}
					});
				
				
					for(int x = 0; x < readClass.customerAmount; x++) {
					
						String kundNamn = Read.data[x] [0] [0];
						
						//Kundens namn st�r alltid f�rst p� mapparna. Detta g�r att ordningen i excel-filen �r den samma som outputen.
						for(int n = 0; n < directories.length; n++ ) {
							if(directories[n].startsWith(kundNamn)) {
								ordning [x] = directories[n];
						}
					
					}
							//Skriver ut ordningen f�r mapparna i konsolen.
					
					
					}
					
					//Loopar kund 
					for(int z = 0; z < readClass.customerAmount; z++) 
					{
						
						//Anropen methoden i Write-klassen
						Write.write(z, inputSourceBox, ordning[z]);
						//rensar f�rg�ende data
						Read.data[z] = null;
				
					}	
				
					//klar meddelande
					run.setEnabled(false);
					klar.setVisible(true);
				
				
		
					
				} catch (IOException e) {
					e.printStackTrace();
				} catch (EncryptedDocumentException e) {
					e.printStackTrace();
				} catch (InvalidFormatException e) {
					e.printStackTrace();
				}
				
			
			}
		});
		frmExcelTillMappar.getContentPane().add(run);
		
		JLabel lblBoxMapp = new JLabel("V\u00E4lj - Huvudmapp");
		springLayout.putConstraint(SpringLayout.EAST, lblBoxMapp, -364, SpringLayout.EAST, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.WEST, clo_dir, 17, SpringLayout.EAST, lblBoxMapp);
		springLayout.putConstraint(SpringLayout.WEST, lblBoxMapp, 10, SpringLayout.WEST, frmExcelTillMappar.getContentPane());
		lblBoxMapp.setHorizontalAlignment(SwingConstants.RIGHT);
		frmExcelTillMappar.getContentPane().add(lblBoxMapp);
		
		JLabel lblExcelMassadata = new JLabel("V\u00E4lj Excel-filen");
		springLayout.putConstraint(SpringLayout.NORTH, lblExcelMassadata, 211, SpringLayout.NORTH, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.SOUTH, lblBoxMapp, -23, SpringLayout.NORTH, lblExcelMassadata);
		springLayout.putConstraint(SpringLayout.EAST, lblExcelMassadata, -364, SpringLayout.EAST, frmExcelTillMappar.getContentPane());
		springLayout.putConstraint(SpringLayout.WEST, excel_dir, 17, SpringLayout.EAST, lblExcelMassadata);
		lblExcelMassadata.setHorizontalAlignment(SwingConstants.RIGHT);
		springLayout.putConstraint(SpringLayout.WEST, lblExcelMassadata, 10, SpringLayout.WEST, frmExcelTillMappar.getContentPane());
		frmExcelTillMappar.getContentPane().add(lblExcelMassadata);
	}
}
