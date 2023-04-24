import imaplib
import email

# Login to your email account
imap_server = "imap-mail.outlook.com"


imap = imaplib.IMAP4_SSL(imap_server)
imap.login(email_address, password)

imap.select("Inbox")

_, msgnums = imap.search(None, "ALL")

print(msgnums)
