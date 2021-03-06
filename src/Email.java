import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

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
	private JTextField cc;
		

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
		frame.setTitle("Csoportos email k??ld??s");
		
		beolvas = new JButton("C??m lista");
		beolvas.setBounds(106, 11, 89, 23);
		beolvas.addActionListener(new Megnyitas());
		frame.getContentPane().add(beolvas);
		
		JButton kuldes = new JButton("K??ld??s");
		kuldes.setBounds(434, 453, 89, 23);
		kuldes.addActionListener(new Kuldes());
		frame.getContentPane().add(kuldes);
		
		csatol = new JButton("Csatoland??");
		csatol.setBounds(106, 400, 89, 23);
		csatol.addActionListener(new Csatolmany());
		frame.getContentPane().add(csatol);
		
		level = new JTextArea();
		level.setBounds(10, 10, 445, 231);
		frame.getContentPane().add(level);
		
		targy = new JTextField();
		targy.setBounds(106, 115, 256, 20);
		frame.getContentPane().add(targy);
		targy.setColumns(10);
		
		JLabel targymegnevezes = new JLabel("T??rgy:");
		targymegnevezes.setBounds(26, 118, 70, 14);
		frame.getContentPane().add(targymegnevezes);
		
		JLabel cimzettek = new JLabel("C??mzettek");
		cimzettek.setBounds(26, 15, 46, 14);
		frame.getContentPane().add(cimzettek);
		
		JLabel lblNewLabel = new JLabel("Lev??l tartalma:");
		lblNewLabel.setBounds(26, 151, 70, 14);
		frame.getContentPane().add(lblNewLabel);
		
		JLabel csatolmany = new JLabel("Csatolm??ny");
		csatolmany.setBounds(26, 404, 46, 14);
		frame.getContentPane().add(csatolmany);
		
		felado = new JTextField();
		felado.setText("@veas.videoton.hu");
		felado.setBounds(106, 45, 256, 20);
		frame.getContentPane().add(felado);
		felado.setColumns(10);
		
		JLabel felad = new JLabel("Felad??");
		felad.setBounds(26, 48, 46, 14);
		frame.getContentPane().add(felad);
		
		JScrollPane scrollPane = new JScrollPane(level);
		scrollPane.setBounds(106, 146, 445, 231);
		frame.getContentPane().add(scrollPane);
		
		JButton elonezet = new JButton("El??n??zet");
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
		
		JLabel masolat = new JLabel("CC");
		masolat.setBounds(26, 81, 46, 14);
		frame.getContentPane().add(masolat);
		
		cc = new JTextField();
		cc.setText("@veas.videoton.hu");
		cc.setBounds(106, 76, 256, 20);
		frame.getContentPane().add(cc);
		cc.setColumns(10);
		
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
	
			int szamlalo2 = 0;
			
			Logger logger = Logger.getAnonymousLogger();
			
			for(int szamlalo = 1; szamlalo < emailcimek.size(); szamlalo++)
			{	
				//final String username = "kovacs.zoltan@veas.videoton.hu";										
		        //final String password = "*******";																
		        
		        //System.setProperty("mail.smtp.ssl.protocols", "TLSv1.2");
		        Properties props = new Properties(); //new Properties();     System.getProperties();
		        
		        props.put("mail.smtp.host", "172.20.22.254");					//smtp.gmail.com					//172.20.22.254 bels?? levelez??s      //smtp-mail.outlook.com
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
		        Session session = Session.getInstance(props, null);									//session l??trehoz??sa a megadott param??terekkel
		        try 
		        {
		            Message message = new MimeMessage(session);
		            message.setFrom(new InternetAddress(felado.getText()));							//felad?? be??ll??t??sa
		            message.setRecipients(Message.RecipientType.TO,
		                InternetAddress.parse(emailcimek.get(szamlalo)));							//c??mzett be??ll??t??sa
		            message.setRecipients(Message.RecipientType.CC,
			                InternetAddress.parse(cc.getText()));							//c??mzett be??ll??t??sa
		            message.setSubject(targy.getText());											//t??rgy be??ll??t??sa
		           
		            Multipart multipart = new MimeMultipart();										//csatol?? oszt??ly p??ld??nyos??t??sa
	
		            MimeBodyPart attachmentPart = new MimeBodyPart();								//csatolm??ny oszt??ly p??ld??nyos??t??sa
		            MimeBodyPart attachmentPart2 = new MimeBodyPart();								//csatolm??ny oszt??ly p??ld??nyos??t??sa
		            MimeBodyPart attachmentPart3 = new MimeBodyPart();								//csatolm??ny oszt??ly p??ld??nyos??t??sa
	
		            MimeBodyPart textPart = new MimeBodyPart();										//lev??l sz??veg??nyek oszt??ly p??ld??nyos??t??sa
	
		            if(fix2 != null)																//ha csatoltak 2. fix csatolm??nyt is akkor fut le
		            {
		                attachmentPart.attachFile(mappa[szamlalo2]);									//csatolm??ny csatol??sa
		                attachmentPart2.attachFile(fix1);											//fix csatolm??ny csatol??sa
		                attachmentPart3.attachFile(fix2);											//fix csatolm??ny csatol??sa
		                textPart.setText(level.getText());											//lev??l tartalm??nak csatol??sa
		                multipart.addBodyPart(textPart);											//csatolm??ny oszt??ly 
		                multipart.addBodyPart(attachmentPart);
		                multipart.addBodyPart(attachmentPart2);
		                multipart.addBodyPart(attachmentPart3);
		            }
		            else if(fix1 != null)																			//amennyiben csak 1 fix van csatolva
		            {
		            	attachmentPart.attachFile(mappa[szamlalo2]);
		                attachmentPart2.attachFile(fix1);
		                textPart.setText(level.getText());
		                multipart.addBodyPart(textPart);
		                multipart.addBodyPart(attachmentPart);
		                multipart.addBodyPart(attachmentPart2);
		            }
		            else
		            {
		            	attachmentPart.attachFile(mappa[szamlalo2]);
		                textPart.setText(level.getText());
		                multipart.addBodyPart(textPart);
		                multipart.addBodyPart(attachmentPart);
		            }
		        
		           
		            if(szamlalo < emailcimek.size())
		            {
		            	szamlalo++;
		            	szamlalo2++;
		            }
		            else
		            {
		            	szamlalo2++;
		            }
		            message.setContent(multipart);													//message ??zenethez mindent hozz??ad
		            
		            Transport.send(message);														//lev??l k??ld??se
	
		            System.out.println("Done");														//ki??rja, ha lefutott minden rendben
		        
		        }
		        catch (Exception e1) 
		        {
		        	String hibauzenet = e1.toString();
		        	logger.log(Level.SEVERE, "an exception was thrown", e1);
	                JOptionPane.showMessageDialog(null, emailcimek.get(szamlalo) + hibauzenet, "Hiba ??zenet", 2);
	                FileWriter fstream;
					try {
						fstream = new FileWriter("c:\\Users\\jenei.erika\\Desktop\\hiba.txt");
						BufferedWriter out=new BufferedWriter(fstream);
		                out.write(e1.toString()+"\n");
		                out.write(emailcimek.get(szamlalo) +"  "+  mappa[szamlalo2].getName());
		                out.close();
					} catch (IOException e2) {
						// TODO Auto-generated catch block
						String hibauzenet2 = e2.toString();
						JOptionPane.showMessageDialog(null, hibauzenet2, "Hiba ??zenet", 2);
						e2.printStackTrace();
					}
	            
	                break;
		        }
		        /*
				catch (IOException e1) 																//Exception kiv??telek eset??n t??rt??nik
	            {
	                String hibauzenet = e1.toString();  											//hiba??zenet string?? alak??t??sa
	                
	                JOptionPane.showMessageDialog(null, emailcimek.get(szamlalo) + hibauzenet, "Hiba ??zenet", 2);				//hiba??zenet kiirat??sa egy kis ablakba
	                break;
	            }
		        catch (MessagingException e1) 
		        {
		        	String hibauzenet = e1.toString();
		        	logger.log(Level.SEVERE, "an exception was thrown", e1);
	                JOptionPane.showMessageDialog(null, emailcimek.get(szamlalo) + hibauzenet, "Hiba ??zenet", 2);
	                break;
		        }
		        catch (NullPointerException e1) 
	            {
	                String hibauzenet = e1.toString();  
	                JOptionPane.showMessageDialog(null, emailcimek.get(szamlalo) + hibauzenet, "Hiba ??zenet", 2);
	                break;
	            }
		        catch (ArrayIndexOutOfBoundsException e1) 
	            {
	                String hibauzenet = e1.toString();  
	                JOptionPane.showMessageDialog(null, emailcimek.get(szamlalo) + hibauzenet, "Hiba ??zenet", 2);
	                break;
	            }*/
			}			
			JOptionPane.showMessageDialog(null, "K??ld??s k??sz", "T??j??koztat?? ??zenet", 1);			
		}		
	}
	
	private class Megnyitas implements ActionListener															//megnyit?? oszt??ly
	{
		public void actionPerformed(ActionEvent e)
		 {
			try
			{
				if (e.getSource() == beolvas) 
				{
					fc.setCurrentDirectory(new java.io.File("z:\\RoHS,Reach, CFSI\\"));
					int returnVal = fc.showOpenDialog(frame);											//f??jl megniyt??s??nak adbalak megnyit
	
					if (returnVal == JFileChooser.APPROVE_OPTION) 					
					{
						File file = fc.getSelectedFile();												//f??jl v??ltoz?? megkpja azt a f??jlt amit kiv??lsztottunk a filechooserrel

		            	FileInputStream fis = new FileInputStream(file);								//file input stream oszt??ly l??trehoz??sa a kiv??lasztott f??jlal
						XSSFWorkbook workbook = new XSSFWorkbook(fis);  								//Excel oszt??ly l??trehoz??sa
		            	XSSFSheet sheet = workbook.getSheetAt(0);										//excel t??bla l??trehoz??sa
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
				JOptionPane.showMessageDialog(null, "Olvas??si hiba t??rt??nt", "Hiba??zenet", 2);
			}/*
			for(int szamlalo = 3; szamlalo < emailcimek.size(); szamlalo++)
			{
				System.out.println(emailcimek.get(szamlalo));
				szamlalo ++;
			}
			for(int szamlalo = 2; szamlalo < emailcimek.size(); szamlalo++)
			{
				System.out.println(emailcimek.get(szamlalo));
				szamlalo ++;
			}*/
		 }		
	}
	
	private class Csatolmany implements ActionListener															//megniyt?? oszt??ly
	{
		public void actionPerformed(ActionEvent e)
		 {
			if (e.getSource() == csatol) 
			{
				fc2.setCurrentDirectory(new java.io.File("z:\\RoHS,Reach, CFSI\\"));
				fc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);							//csak k??nyvt??rakat nyit meg
				fc2.setAcceptAllFileFilterUsed(false);												//kikapcsolja a f??jlok l??that??s??g??t
				int returnVal = fc2.showOpenDialog(frame);											//f??jl megniyt??s??nak adbalak megnyit	
 
				if (returnVal == JFileChooser.APPROVE_OPTION) 
				{
					csatoltfile = fc2.getSelectedFile();											//file oszt??lynak odaadja a kiv??laszott mapp??t
					mappa = csatoltfile.listFiles();												//kilist??zza ??s egy t??mbnek adja a mappa elemeit
				}				
			}
			/*
			for(int szamlalo = 0; szamlalo < mappa.length; szamlalo++)
			{
				System.out.println(mappa[szamlalo]);
			}
			*/ 
		 }		
	}
	
	private class Elonezet implements ActionListener														//el??n??zet osz??ly
	{
		public void actionPerformed(ActionEvent e) 
		{
			parbeszed();																			//parbesz??d met??dus megh??v??sa
			
		}
	}
	
	void parbeszed()																				//met??dus, ami megmutatja mik vannak csatolva, milyen email c??mekhez
	{
		JFrame ablak = new JFrame();																//??j ablak l??trehoz??sa
		ablak.setBounds(200, 200, 1000, 400);
		ablak.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
		ablak.getContentPane().setLayout(null);
		ablak.setTitle("Csatol??si el??n??zet");
		
		DefaultListModel<String> model = new DefaultListModel<String>();	
		JList<String> lista = new JList<String>();
		//lista.setBounds(100, 50, 800, 200);
		lista.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);								//1x-es kijel??l??s be??ll??t??sa
		lista.setModel(model);																		//listamodell be??ll??t??sa
		ablak.getContentPane().add(lista);
		
		JScrollPane scrollPane2 = new JScrollPane(lista);											//scrollozhat?? ablak l??trehoz??sa a Jlistb??l
		scrollPane2.setBounds(100, 30, 800, 300);
		ablak.getContentPane().add(scrollPane2);
		int szamlalo;
		int szamlalo2 = 0;
		
		try
		{	
			if(fix2 != null)
	        {
				for(szamlalo = 1; szamlalo2 < mappa.length; szamlalo++)
				{
					model.addElement(emailcimek.get(szamlalo) + "  --  " + mappa[szamlalo2].getName() + "  --  " + fix1.getName() + "  --  " + fix2.getName());
					if(szamlalo < emailcimek.size())
		            {
		            	szamlalo++;
		            }
					szamlalo2++;
				} 
	        }
			else if(fix1 != null)
			{
				for(szamlalo = 1; szamlalo2 < mappa.length; szamlalo++)
				{
					model.addElement(emailcimek.get(szamlalo) + "  --  " + mappa[szamlalo2].getName() + "  --  " + fix1.getName());
					if(szamlalo < emailcimek.size())
		            {
		            	szamlalo++;
		            }
					szamlalo2++;
				}
			}
			else
			{
				for(szamlalo = 1; szamlalo2 < mappa.length; szamlalo++)
				{
					model.addElement(emailcimek.get(szamlalo) + "  --  " + mappa[szamlalo2].getName());
					if(szamlalo < emailcimek.size())
		            {
		            	szamlalo++;
		            }
					szamlalo2++;
				}
			}
		}
		catch(Exception e1) 																//Exception kiv??telek eset??n t??rt??nik
        {
            String hibauzenet = e1.toString();  											//hiba??zenet string?? alak??t??sa
            JOptionPane.showMessageDialog(null, hibauzenet, "Hiba ??zenet", 2);				//hiba??zenet kiirat??sa egy kis ablakba
            ablak.setVisible(true);
        }
		
		ablak.setVisible(true);
	}
	
	private class FixCsatolmany1 implements ActionListener															//megniyt?? oszt??ly
	{
		public void actionPerformed(ActionEvent e)
		 {
			if (e.getSource() == fix1csatol) 
			{
				fc3.setCurrentDirectory(new java.io.File("z:\\RoHS,Reach, CFSI\\"));
				int returnVal = fc3.showOpenDialog(frame);													//f??jl megniyt??s??nak adbalak megnyit
				
				if (returnVal == JFileChooser.APPROVE_OPTION) 
				{
					fix1 = fc3.getSelectedFile();		
				}
			}

		 }		
	}
	
	private class FixCsatolmany2 implements ActionListener															//megniyt?? oszt??ly
	{
		public void actionPerformed(ActionEvent e)
		 {
			if (e.getSource() == fix2csatol) 
			{
				fc4.setCurrentDirectory(new java.io.File("z:\\RoHS,Reach, CFSI\\"));
				int returnVal = fc4.showOpenDialog(frame);											//f??jl megniyt??s??nak adbalak megnyit
				
				if (returnVal == JFileChooser.APPROVE_OPTION) 
				{
					fix2 = fc4.getSelectedFile();		
				}
			}
		 }		
	}
	private static class __Tmp {
		private static void __tmp() {
			  javax.swing.JPanel __wbp_panel = new javax.swing.JPanel();
		}
	}
}
