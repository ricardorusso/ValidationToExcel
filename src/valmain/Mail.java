package valmain;

import java.io.IOException;
import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.NoSuchProviderException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Store;

public class Mail {

	public static void check(String host, String storeType, String user,
		      String password) 
		   {
		      try {
		    	  final String fromEmail = "pt.aubay@gmail.com"; //requires valid gmail id
					final String password2 = "aubay123"; // correct password for gmail id
		      //create properties field
		      Properties properties = new Properties();

		      properties.put("mail.smtp.host", "smtp.gmail.com"); //SMTP Host
		      properties.put("mail.smtp.socketFactory.port", "465"); //SSL Port
		      properties.put("mail.smtp.socketFactory.class",
						"javax.net.ssl.SSLSocketFactory"); //SSL Factory Class
		      properties.put("mail.smtp.auth", "true"); //Enabling SMTP Authentication
				properties.put("mail.smtp.port", "465"); //SMTP Port

				Authenticator auth = new Authenticator() {
					//override the getPasswordAuthentication method
					protected PasswordAuthentication getPasswordAuthentication() {
						return new PasswordAuthentication(fromEmail, password2);
					}
				};
		      Session emailSession = Session.getDefaultInstance(properties,auth);
		  
		      //create the POP3 store object and connect with the pop server
		      Store store = emailSession.getStore("pop3s");

		      store.connect(host, user, password);
		      System.out.println("Conected");
		      //create the folder object and open it
		      Folder emailFolder = store.getFolder("INBOX");
		      emailFolder.open(Folder.READ_ONLY);

		      // retrieve the messages from the folder in an array and print it
		      Message[] messages = emailFolder.getMessages();
		      System.out.println("messages.length---" + messages.length);

		      for (int i = 0, n = messages.length; i < n; i++) {
		         Message message = messages[i];
		         System.out.println("---------------------------------");
		         System.out.println("Email Number " + (i + 1));
		         System.out.println("Subject: " + message.getSubject());
		         System.out.println("From: " + message.getFrom()[0]);
		         System.out.println("Text: " + message.getContent().toString());

		      }

		      //close the store and folder objects
		      emailFolder.close(false);
		      store.close();

		      } catch (NoSuchProviderException e) {
		         e.printStackTrace();
		      } catch (MessagingException e) {
		         e.printStackTrace();
		      } catch (Exception e) {
		         e.printStackTrace();
		      }
		   }

		   public static void main(String[] args) {

		      String host = "smtp.gmail.com";// change accordingly
		      String mailStoreType = "pop";
		      final String username = "pt.aubay@gmail.com"; //requires valid gmail id
				final String password = "aubay123"; // correct password for gmail id

		      check(host, mailStoreType, username, password);

		   }
}
