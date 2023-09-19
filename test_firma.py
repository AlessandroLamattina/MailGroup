import os
import win32com.client as win32

# Crea un oggetto Outlook
outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace("MAPI")

# Crea un'email
mail = outlook.CreateItem(0)  # 0 indica un'email standard

# Imposta i destinatari e l'oggetto dell'email
mail.To = 'alamattinalarocca@deloitte.it'  # Imposta l'indirizzo email del destinatario
mail.Subject = 'Oggetto dell\'email'  # Imposta l'oggetto dell'email

# Specifica il percorso del file Word da incorporare nel corpo dell'email
file_word = r"C:\Users\alamattinalarocca\Downloads\Corso OIC 34 - DELOITTE_TEMPLATE 2.docx"
file_html=r"C:\Users\alamattinalarocca\Downloads\Corso OIC 34 - DELOITTE_TEMPLATE 2.html"
# Verifica se il file esiste prima di procedere
if os.path.exists(file_word):
    with open(file_html, 'rb') as file:
        file_content = file.read()

    html_body = f"<html><body><p>Contenuto del documento Word:</p><br><p>{file_content}</p></body></html>"
    mail.HTMLBody = html_body
    mail.Send()
else:
    print(f"Il file '{file_word}' non esiste. Impossibile inviare l'email.")
