import os
import win32com.client as win32

outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace("MAPI")
mail = outlook.CreateItem(0)  # 0 indica un'email standard

mail.To = 'alamattinalarocca@deloitte.it'  # Imposta l'indirizzo email del destinatario
mail.Subject = 'Oggetto dell\'email'  # Imposta l'oggetto dell'email
mail.Body = 'Corpo dell\'email'  # Imposta il corpo dell'email

signature_path = os.path.expandvars('%APPDATA%\\Microsoft\\Signatures\\')
signature_file = os.listdir(signature_path)[0]
# Leggi il contenuto del file della firma
with open(signature_path + signature_file, 'r') as file:
    signature = file.read()
# Aggiungi la firma al corpo dell'email
mail.HTMLBody += '<br><br>' + signature
mail.send()