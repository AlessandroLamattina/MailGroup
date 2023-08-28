import pyodbc
import win32com.client as win32

# Connessione al database MDB
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\test.mdb;')
cursor = conn.cursor()

# Recupero degli indirizzi email e delle società dal database
cursor.execute("SELECT numero, email, azienda, sesso, titolo, cognome FROM Tabella1")#TestEmail
outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace("MAPI")
results = cursor.fetchall()

# Raggruppamento degli indirizzi email per dominio aziendale
email_dict = {}

for numero, email, azienda, sesso, titolo, cognome in sorted(results, key=lambda x: x[0]):
    domain = azienda
    if domain not in email_dict:
        email_dict[domain] = []
    email_dict[domain].append([numero, email, sesso, titolo, cognome])

#Invio delle email ai destinatari della stessa azienda


for domain, recipients in email_dict.items():
    message = outlook.CreateItem(0)
    message.Subject = ""
    message.Body = ""
    # Aggiunta dei destinatari
    i=0

    recipients_list = message.Recipients
    for numero, email, sesso, titolo, cognome in recipients:
        recipient_email = namespace.CreateRecipient(email)
        recipient_email = namespace.CreateRecipient(email)
        recipient_email.Type = 1 if i == 0 else 3  # Imposta il tipo come "To" per il primo destinatario, "CC" per gli altri

        recipients_list.Add(recipient_email)
        if i==0:
            message.Body= f"{sesso} {titolo} {cognome}\n" \
                          f"ti invio questa offerta per la tua azienda {domain}\n" \
                          "" \
                          "Cordiali Saluti\n" \
                          "Alessandro\n" \
                          "" \

            i+=1
    recipients_list.ResolveAll()
    try:
        # Invio dell'email
        message.Send()
        print("Email inviata con successo a", domain)
    except Exception as e:
        print("Errore durante l'invio dell'email a", domain)
        print(str(e))

# Chiusura della connessione al database
conn.close()

