import email
import smtplib
import imaplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

GMAIL_SMTP = "smtp.gmail.com"
GMAIL_IMAP = "imap.gmail.com"

class Outlook():
    def __init__(self, e_mail, password):
        self.e_mail = e_mail
        self.password = password

    def send_message(self, recipients, subject, message):
        self.recipients = recipients
        self.subject = subject
        self.message = message
        msg = MIMEMultipart()
        msg['From'] = e_mail
        msg['To'] = ', '.join(self.recipients)
        msg['Subject'] = self.subject
        msg.attach(MIMEText(self.message))

        ms = smtplib.SMTP(GMAIL_SMTP, 587)
        ms.ehlo()
        ms.starttls()
        ms.ehlo()

        ms.login(self.e_mail, self.password)
        ms.sendmail(self.e_mail, ms, msg.as_string())
        ms.quit()

    def receive(self, header):
        self.header = header
        mail = imaplib.IMAP4_SSL(GMAIL_IMAP)
        mail.login(self.e_mail, self.password)
        mail.list()
        mail.select("inbox")
        criterion = '(HEADER Subject "%s")' % self.header if self.header else 'ALL'
        result, data = mail.uid('search', None, criterion)
        assert data[0], 'There are no letters with current header'
        latest_email_uid = data[0].split()[-1]
        result, data = mail.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = data[0][1]
        email_message = email.message_from_string(raw_email)
        mail.logout()


if __name__ == "__main__":
    e_mail = 'login@gmail.com'
    password = 'qwerty'
    subject = 'Subject'
    recipients = ['vasya@email.com', 'petya@email.com']
    message = 'Message'
    header = None

    pochta = Outlook(e_mail, password)
    pochta.send_message(recipients, subject, message)
    pochta.receive(header)
