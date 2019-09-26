package valtoexcel;

import java.io.File;
import java.io.InputStream;
import java.time.LocalTime;
import java.util.List;
import java.util.Properties;
import java.util.logging.Logger;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Folder;
import javax.mail.Message.RecipientType;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.formula.functions.T;


public class Mail {
	

	private static final Logger logger = Logger.getGlobal();
	private String subject, imageDir, recipients, bodyText;

	private String saluts;
	
	private File anexoDir;
	private List<?> listResume;
	
	public void mailGenarator() 
	{
		try {
			logger.info("Generating Draft Email");
			
			File imageFile = new File(this.getImageDir());
			File anexoDirEmail = this.getAnexoDir();
			String subjectEmail = this.getSubject();

			InputStream iS = Mail.class.getResourceAsStream("/mail.properties");
			Properties props = new Properties();
			Properties authProp =  new Properties();
			InputStream iSau = Mail.class.getResourceAsStream("/auth.properties");
			authProp.load(iSau);
			props.setProperty("mail.imap.ssl.enable", "true");
			props.setProperty("mail.imap.ssl.trust", "*");
			props.load(iS);

			
			Authenticator auth = new Authenticator() {
				
				@Override
				protected PasswordAuthentication getPasswordAuthentication() {
					return new PasswordAuthentication(authProp.getProperty("fromEmail"), authProp.getProperty("pass"));
				}
			};
			Session emailSession = Session.getDefaultInstance(props,auth);

			//create the POP3 store object and connect with the pop server
			Store store = emailSession.getStore("imaps");
			System.out.println(store.isConnected());
			store.connect("outlook.office365.com", authProp.getProperty("fromEmail"), authProp.getProperty("pass") );
			logger.info("Conected do Outlook email");

			
			MimeMessage msg = new MimeMessage(emailSession);
			msg.setFlag(javax.mail.Flags.Flag.DRAFT,true);
			
			
			Properties mailAdresses =   new Properties();
			mailAdresses.load(Mail.class.getResourceAsStream("/emails.properties"));
			
			msg.addRecipients(RecipientType.TO, mailRecepients(mailAdresses) );
			
			msg.setSubject(subjectEmail);

			
			MimeMultipart multipart = new MimeMultipart();

			
			BodyPart messageBodyPart = new MimeBodyPart();
			String htmlText = getSaluts()
					+getBodyText()
					+ "<img src=\"cid:image\">";
			messageBodyPart.setContent(htmlText, "text/html");
		
			multipart.addBodyPart(messageBodyPart);

		
			messageBodyPart = new MimeBodyPart();

			DataSource fds = new FileDataSource(
					imageFile);

			messageBodyPart.setDataHandler(new DataHandler(fds));
			messageBodyPart.setHeader("Content-ID", "<image>");

 
			multipart.addBodyPart(messageBodyPart);

			//anexo
			DataSource dSa = new FileDataSource(
					anexoDirEmail);
			BodyPart anexo = new MimeBodyPart(); 
			anexo.setDataHandler(new DataHandler(dSa));
			anexo.setFileName(anexoDirEmail.getName());
			multipart.addBodyPart(anexo);
			// put everything together
			msg.setContent(multipart);

			saveDraft(store, msg);

		} catch (Exception e ) {
			e.printStackTrace();
		} 
	}

	private Address[] mailRecepients(Properties mailAdresses) throws AddressException {
		
		String adressesProp =  mailAdresses.getProperty("emails");
		String [] arrEmail = adressesProp.split(";");
		Address[] adresses = new Address[arrEmail.length];
		for (int i = 0; i < arrEmail.length; i++) {
			adresses[i] = new InternetAddress(arrEmail[i]);
		}
		
		return adresses;
		
	}

	private static void saveDraft(Store store, MimeMessage msg) throws MessagingException {
		
		try {
			MimeMessage [] draftmessagems = {
					msg	  
			};

			Folder  emailFolder = store.getFolder("drafts");
			emailFolder.open(Folder.READ_WRITE);
			emailFolder.appendMessages(draftmessagems);
			logger.info("Draft Saved ");

			
			
		} catch (Exception e) {
			
			e.printStackTrace();
		}finally {
			
			
			store.close();
		}
	}


	public String getSubject() {
		return subject;
	}

	public void setSubject(String subject) {
		this.subject = subject;
	}

	public File getAnexoDir() {
		return anexoDir;
	}

	public void setAnexoDir(File anexoDir) {
		this.anexoDir = anexoDir;
	}

	public String getImageDir() {
		return imageDir;
	}

	public void setImageDir(String imageDir) {
		this.imageDir = imageDir;
	}

	public String getRecipients() {
		return recipients;
	}

	public void setRecipients(String recipients) {
		this.recipients = recipients;
	}

	public String getBodyText() {
		
		return bodyText;
	}


	public void setBodyTextForHtml(List<?> listResume2) {
		final String LI = "<li>";
		final String LIE = "</li>";
		
		StringBuilder strB =  new StringBuilder();
		for (Object t : listResume2) {
			strB.append(LI + t.toString() +LIE+"\n");
		}
		
		setBodyText(strB.toString());
		
	}
	
	@SuppressWarnings("unchecked")
	public List<T> getListResume() {
		return (List<T>) listResume;
	}
	public void setListResume(List<T> listResume) {
		this.listResume = listResume;
	}
	public void setBodyText(String bodyText) {
		this.bodyText = bodyText;
	}
	
	
	public Mail(String subject, File anexoDir, String imageDir, List<?> listFinalResume) {
		super();
		this.subject = subject;
		this.anexoDir = anexoDir;
		this.imageDir = imageDir;
		this.listResume = listFinalResume;
		if(!listResume.isEmpty()) {
			setBodyTextForHtml(listResume);
		}
		
	}
	public Mail() {	}

	public String getSaluts() {
		return saluts;
	}

	public void setSaluts(String saluts) {
		this.saluts = saluts;
	}
	
	
	{
		
		LocalTime time = LocalTime.now();
		int hour = time.getHour();
		String saudations ="";
		
		if(hour >= 12 && hour <=19) {
			saudations ="Boa tarde.";
		}else if (hour >19 ) {
			saudations ="Boa noite.";
		}else {
			saudations ="Bom dia.";
		}
		
			String finalSalusts ="<p>"+ (!saudations.equals("") ? saudations:"Boas") +"</p>\r\n" + 
					"    <p>Em anexo seguem os resultados das monitorizações:</p>";

		setSaluts(finalSalusts);
		
	}
}
