import java.util.Properties;
import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;
import java.io.File;
import java.io.FileNotFoundException;

public class AutoMail {

    public static void main(String[] args) {
        // Recipient's email ID
        String to = "arjun@arabyads.com";

        // Sender's email ID
        String from = "abc@gmail.com";
        final String username = "xyz@gmail.com"; // your Gmail username
        final String appPassword = "abcdefghijklmno"; // your Gmail app password

        // Assuming you are sending email through Gmail SMTP server
        String host = "smtp.gmail.com";

        // Setup mail server properties
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", "587");
        props.put("mail.smtp.ssl.trust", "smtp.gmail.com");

        // Get the Session object
        Session session = Session.getInstance(props,
            new javax.mail.Authenticator() {
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(username, appPassword);
                }
            });

        try {
            // Create a default MimeMessage object
            Message message = new MimeMessage(session);

            // Set From: header field of the header
            message.setFrom(new InternetAddress(from));

            // Set To: header field of the header
            message.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(to));

            // Set Subject: header field
            message.setSubject("Daily Report with Excel Attachment");

            // Create the message part
            BodyPart messageBodyPart = new MimeBodyPart();

            // Now set the actual message
            messageBodyPart.setText("Please find attached the daily report.");

            // Create a multipar message
            Multipart multipart = new MimeMultipart();

            // Set text message part
            multipart.addBodyPart(messageBodyPart);

            // Part two is attachment
            messageBodyPart = new MimeBodyPart();
            String filename = "D:/Automation Testing/Automail.xlsx"; // Update this path to your file
            File file = new File(filename);

            // Check if the file exists
            if (!file.exists()) {
                throw new FileNotFoundException("File not found: " + filename);
            }

            DataSource source = new FileDataSource(filename);
            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName(file.getName());
            multipart.addBodyPart(messageBodyPart);

            // Add the note as an inline comment below the attachment
            // Note: this is an automatic mail, please do not reply.
            // You can also append this note to the message body instead
            // messageBodyPart = new MimeBodyPart();
            //messageBodyPart.setText("Note: this is an automatic mail, please do not reply.");
            // multipart.addBodyPart(messageBodyPart);

            // Send the complete message parts
            message.setContent(multipart);

            // Send message
            Transport.send(message);

            System.out.println("Sent message successfully with Excel attachment...");
        } catch (FileNotFoundException e) {
            System.err.println("Error: " + e.getMessage());
        } catch (MessagingException e) {
            throw new RuntimeException(e);
        }
    }
}
