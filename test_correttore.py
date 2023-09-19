import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configurazioni per l'account Outlook/Hotmail
email_address = "alamattinalarocca@deloitte.it"
password = "Silvia1996.."  # Usa la password di app se l'autenticazione a due fattori è abilitata

# Destinatario dell'email
to_email = "alamattinalarocca@deloitte.it"

# Creazione del messaggio
msg = MIMEMultipart()
msg["From"] = email_address
msg["To"] = to_email
msg["Subject"] = "Oggetto del messaggio"

# Leggi il contenuto del file Word
word_file_path = r"C:\Users\alamattinalarocca\Downloads\Corso OIC 34 - DELOITTE_TEMPLATE 2.docx"
with open(word_file_path, "rb") as file:
    word_content = file.read()

# Aggiungi il contenuto del file Word al corpo del messaggio
attachment = MIMEApplication(word_content, _subtype="docx")
attachment.add_header("Content-Disposition", f"attachment; filename={word_file_path}")
msg.attach(attachment)

# Connessione al server SMTP di Outlook/Hotmail
smtp_server = "smtp.live.com"
smtp_port = 587
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()  # Avvia la connessione TLS (crittografata)

# Effettua l'accesso all'account Outlook/Hotmail
server.login(email_address, password)

# Invia l'email
server.sendmail(email_address, to_email, msg.as_string())

# Chiudi la connessione SMTP
server.quit()
# Connessione al server SMTP di Outlook/Hotmail
smtp_server = "smtp.live.com"
smtp_port = 587
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()  # Avvia la connessione TLS (crittografata)

# Effettua l'accesso all'account Outlook/Hotmail
server.login(email_address, password)

# Invia l'email
server.sendmail(email_address, to_email, msg.as_string())

# Chiudi la connessione SMTP
server.quit()

print("Email inviata con successo!")
