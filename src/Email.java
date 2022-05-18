import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Properties;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.DefaultListModel;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.JTextArea;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JScrollPane;

public class Email 
{

	private JFrame frame;
	private JFileChooser fc;
	private JFileChooser fc2;
	private JFileChooser fc3;
	private JFileChooser fc4;
	private JButton beolvas;
	private JButton csatol;
	private JButton fix1csatol;
	private JButton fix2csatol;
	private ArrayList<String> emailcimek;
	private File csatoltfile;
	private File fix1;
	private File fix2;
	private File[] mappa;
	private JTextArea level;
	private JTextField targy;
	private JTextField felado;
		

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) 
	{
		EventQueue.invokeLater(new Runnable() 
		{
			public void run() 
			{
				try 
				{
					Email window = new Email();
					window.frame.setVisible(true);
				} 
				catch (Exception e) 
				{
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public Email() 
	{
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() 
	{
		frame = new JFrame();
		frame.setBounds(100, 100, 595, 526);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		frame.setTitle("Csoportos email küldés");
		
		beolvas = new JButton("Cím lista");
		beolvas.setBounds(106, 25, 89, 23);
		beolvas.addActionListener(new Megnyitas());
		frame.getContentPane().add(beolvas);
		
		JButton kuldes = new JButton("Küldés");
		kuldes.setBounds(434, 453, 89, 23);
		kuldes.addActionListener(new Kuldes());
		frame.getContentPane().add(kuldes);
		
		csatol = new JButton("Csatolandó");
		csatol.setBounds(106, 400, 89, 23);
		csatol.addActionListener(new Csatolmany());
		frame.getContentPane().add(csatol);
		
		level = new JTextArea();
		level.setBounds(106, 146, 445, 231);
		frame.getContentPane().add(level);
		
		targy = new JTextField();
		targy.setBounds(106, 115, 256, 20);
		frame.getContentPane().add(targy);
		targy.setColumns(10);
		
		JLabel targymegnevezes = new JLabel("Tárgy:");
		targymegnevezes.setBounds(26, 118, 70, 14);
		frame.getContentPane().add(targymegnevezes);
		
		JLabel cimzettek = new JLabel("Címzettek");
		cimzettek.setBounds(26, 29, 46, 14);
		frame.getContentPane().add(cimzettek);
		
		JLabel lblNewLabel = new JLabel("Levél tartalma:");
		lblNewLabel.setBounds(26, 151, 70, 14);
		frame.getContentPane().add(lblNewLabel);
		
		JLabel csatolmany = new JLabel("Csatolmány");
		csatolmany.setBounds(26, 404, 46, 14);
		frame.getContentPane().add(csatolmany);
		
		felado = new JTextField();
		felado.setBounds(106, 74, 256, 20);
		frame.getContentPane().add(felado);
		felado.setColumns(10);
		
		JLabel felad = new JLabel("Feladó");
		felad.setBounds(26, 77, 46, 14);
		frame.getContentPane().add(felad);
		
		JScrollPane scrollPane = new JScrollPane(level);
		scrollPane.setBounds(106, 146, 445, 231);
		frame.getContentPane().add(scrollPane);
		
		JButton elonezet = new JButton("Előnézet");
		elonezet.setBounds(434, 400, 89, 23);
		elonezet.addActionListener(new Elonezet());
		frame.getContentPane().add(elonezet);
		
		fix1csatol = new JButton("Fix1");
		fix1csatol.setBounds(205, 400, 89, 23);
		fix1csatol.addActionListener(new FixCsatolmany1());
		frame.getContentPane().add(fix1csatol);
		
		fix2csatol = new JButton("Fix2");
		fix2csatol.setBounds(304, 400, 89, 23);
		fix2csatol.addActionListener(new FixCsatolmany2());
		frame.getContentPane().add(fix2csatol);
		
		emailcimek = new ArrayList<String>();
		
		fc = new JFileChooser();
		fc2 = new JFileChooser();
		fc3 = new JFileChooser();
		fc4 = new JFileChooser();
	}
	
	private class Kuldes implements ActionListener
	{
		public void actionPerformed(ActionEvent e) 
		{
			try 
	        {
			mappa = csatoltfile.listFiles();												//csatolt mappa tartalmának kilistázása
	        }
			 catch (NullPointerException e1) 
            {
                String hibauzenet = e1.toString();  										//hibaüzenet Stringé alakítása
                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);			//hibaüzenet kiiratása üzenetablakba
            }
			for(int szamlalo = 0; szamlalo < emailcimek.size(); szamlalo++)
			{	
				//final String username = "kovacs.zoltan@veas.videoton.hu";										
		        //final String password = "*******";																
		        
		        //System.setProperty("mail.smtp.ssl.protocols", "TLSv1.2");
		        Properties props = new Properties(); //new Properties();     System.getProperties();
		        
		        props.put("mail.smtp.host", "172.20.22.254");					//smtp.gmail.com					//172.20.22.254 belső levelezés      //smtp-mail.outlook.com
		        props.put("mail.smtp.port", "25");										//587 TLS		//465  SSL			//25 Outlook							//587
		        //props.put("mail.smtp.auth", "true");
		        //props.put("mail.smtp.starttls.enable", "true");
		        //props.put("mail.smtp.starttls.required", "true");
		        //props.put("mail.smtp.ssl.trust", "smtp.gmail.com");	//
		        //props.put("mail.smtp.ssl.enable", "true");				//setProperty
		        //props.put("mail.smtp.ssl.protocols", "TLSv1.2");
		        //props.put("mail.smtp.ssl.enable", "true");
		        
		        /*
		        Authenticator auth = new javax.mail.Authenticator() {
		              protected PasswordAuthentication getPasswordAuthentication() {
		                  return new PasswordAuthentication(username, password);
		              }
		            };
		        */
		        Session session = Session.getInstance(props, null);									//session létrehozűsa a megadott paraméterekkel
		        try 
		        {
		            Message message = new MimeMessage(session);
		            message.setFrom(new InternetAddress(felado.getText()));							//feladó beállítása
		            message.setRecipients(Message.RecipientType.TO,
		                InternetAddress.parse(emailcimek.get(szamlalo)));							//címzett beállítása
		            message.setSubject(targy.getText());											//tárgy beállítása
		           
		            Multipart multipart = new MimeMultipart();										//csatoló osztály példányosítása
	
		            MimeBodyPart attachmentPart = new MimeBodyPart();								//csatolmány osztály példányosítása
		            MimeBodyPart attachmentPart2 = new MimeBodyPart();								//csatolmány osztály példányosítása
		            MimeBodyPart attachmentPart3 = new MimeBodyPart();								//csatolmány osztály példányosítása
	
		            MimeBodyPart textPart = new MimeBodyPart();										//levél szövegények osztály példányosítása
	
		            if(fix2 != null)																//ha csatoltak 2. fix csatolmányt is akkor fut le
		            {
		                attachmentPart.attachFile(mappa[szamlalo]);									//csatolmány csatolása
		                attachmentPart2.attachFile(fix1);											//fix csatolmány csatolása
		                attachmentPart3.attachFile(fix2);											//fix csatolmány csatolása
		                textPart.setText(level.getText());											//levél tartalmának csatolása
		                multipart.addBodyPart(textPart);											//csatolmány osztály 
		                multipart.addBodyPart(attachmentPart);
		                multipart.addBodyPart(attachmentPart2);
		                multipart.addBodyPart(attachmentPart3);
		            }
		            else																			//amennyiben csak 1 fix van csatolva
		            {
		            	attachmentPart.attachFile(mappa[szamlalo]);
		                attachmentPart2.attachFile(fix1);
		                textPart.setText(level.getText());
		                multipart.addBodyPart(textPart);
		                multipart.addBodyPart(attachmentPart);
		                multipart.addBodyPart(attachmentPart2);
		            }
		        
		           
	
		            message.setContent(multipart);													//message üzenethez mindent hozzáad
		            
		            Transport.send(message);														//levél küldése
	
		            System.out.println("Done");														//kiírja, ha lefutott minden rendben
		        
		        }
				catch (IOException e1) 																//Exception kivételek esetén történik
	            {
	                String hibauzenet = e1.toString();  											//hibaüzenet stringé alakítása
	                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);				//hibaüzenet kiiratása egy kis ablakba
	            }
		        catch (MessagingException e1) 
		        {
		        	String hibauzenet = e1.toString();  
	                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);
		        }
		        catch (NullPointerException e1) 
	            {
	                String hibauzenet = e1.toString();  
	                JOptionPane.showMessageDialog(null, hibauzenet, "Hiba üzenet", 2);
	            }
			}			
			JOptionPane.showMessageDialog(null, "Küldés kész", "Tájékoztató üzenet", 1);			
		}		
	}
	
	private class Megnyitas implements ActionListener															//megnyitó osztály
	{
		public void actionPerformed(ActionEvent e)
		 {
			try
			{
				if (e.getSource() == beolvas) 
				{

					int returnVal = fc.showOpenDialog(frame);											//fájl megniytásának adbalak megnyit
	
					if (returnVal == JFileChooser.APPROVE_OPTION) 					
					{
						File file = fc.getSelectedFile();												//fájl változó megkpja azt a fájlt amit kiválsztottunk a filechooserrel

		            	FileInputStream fis = new FileInputStream(file);								//file input stream osztály létrehozása a kiválasztott fájlal
						XSSFWorkbook workbook = new XSSFWorkbook(fis);  								//Excel osztály létrehozása
		            	XSSFSheet sheet = workbook.getSheetAt(0);										//excel tábla létrehozása
		            	Iterator<Row> itr = sheet.iterator();    										//iterating over excel file  
						
		            	while (itr.hasNext())                 
		            	{  
			            	Row row = itr.next();  
			            	Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
			            	while (cellIterator.hasNext())   
			            	{  
			            		Cell cell = cellIterator.next();
			            		emailcimek.add(cell.getStringCellValue());
			            	}  			            	 
		            	}  
					} 
				}
			}
			catch(IOException e1)
			{
				JOptionPane.showMessageDialog(null, "Olvasási hiba történt", "Hibaüzenet", 2);
			}
		 }		
	}
	
	private class Csatolmany implements ActionListener															//megniytó osztály
	{
		public void actionPerformed(ActionEvent e)
		 {
			if (e.getSource() == csatol) 
			{
				fc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);							//csak könyvtárakat nyit meg
				fc2.setAcceptAllFileFilterUsed(false);												//kikapcsolja a fájlok láthatóságát
				int returnVal = fc2.showOpenDialog(frame);											//fájl megniytásának adbalak megnyit	
 
				if (returnVal == JFileChooser.APPROVE_OPTION) 
				{
					csatoltfile = fc2.getSelectedFile();											//file osztálynak odaadja a kiválaszott mappát
					mappa = csatoltfile.listFiles();												//kilistázza és egy tömbnek adja a mappa elemeit
				}				
			}
		 }		
	}
	
	private class Elonezet implements ActionListener														//előnézet oszály
	{
		public void actionPerformed(ActionEvent e) 
		{
			parbeszed();																			//parbeszéd metódus meghívása
			
		}
	}
	
	void parbeszed()																				//metódus, ami megmutatja mik vannak csatolva, milyen email címekhez
	{
		JFrame ablak = new JFrame();																//új ablak létrehozása
		ablak.setBounds(200, 200, 1000, 400);
		ablak.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
		ablak.getContentPane().setLayout(null);
		ablak.setTitle("Csatolási előnézet");
		
		DefaultListModel<String> model = new DefaultListModel<String>();	
		JList<String> lista = new JList<String>();
		//lista.setBounds(100, 50, 800, 200);
		lista.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);								//1x-es kijelõlés beállítása
		lista.setModel(model);																		//listamodell beállítása
		ablak.getContentPane().add(lista);
		
		JScrollPane scrollPane2 = new JScrollPane(lista);											//scrollozható ablak létrehozása a Jlistből
		scrollPane2.setBounds(100, 30, 800, 300);
		ablak.getContentPane().add(scrollPane2);
		
		if(fix2 != null)
        {
			for(int szamlalo = 0; szamlalo < emailcimek.size(); szamlalo++)
			{
				model.addElement(emailcimek.get(szamlalo) + "  --  " + mappa[szamlalo].getName() + "  --  " + fix1.getName() + "  --  " + fix2.getName());
			} 
        }
		else
		{
			for(int szamlalo = 0; szamlalo < emailcimek.size(); szamlalo++)
			{
				model.addElement(emailcimek.get(szamlalo) + "  --  " + mappa[szamlalo].getName() + "  --  " + fix1.getName());
			}
		}
		
		ablak.setVisible(true);
	}
	
	private class FixCsatolmany1 implements ActionListener															//megniytó osztály
	{
		public void actionPerformed(ActionEvent e)
		 {
			if (e.getSource() == fix1csatol) 
			{
				int returnVal = fc3.showOpenDialog(frame);													//fájl megniytásának adbalak megnyit
				
				if (returnVal == JFileChooser.APPROVE_OPTION) 
				{
					fix1 = fc3.getSelectedFile();		
				}
			}

		 }		
	}
	
	private class FixCsatolmany2 implements ActionListener															//megniytó osztály
	{
		public void actionPerformed(ActionEvent e)
		 {
			if (e.getSource() == fix2csatol) 
			{
				int returnVal = fc4.showOpenDialog(frame);											//fájl megniytásának adbalak megnyit
				
				if (returnVal == JFileChooser.APPROVE_OPTION) 
				{
					fix2 = fc4.getSelectedFile();		
				}
			}
		 }		
	}
	
	
}
