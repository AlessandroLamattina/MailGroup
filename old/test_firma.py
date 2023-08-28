import os
import win32com.client as win32
import base64

outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace("MAPI")
mail = outlook.CreateItem(0)  # 0 indica un'email standard

mail.To = 'lamattinalessandro@gmail.com'  # Imposta l'indirizzo email del destinatario
mail.Subject = 'Oggetto dell\'email'  # Imposta l'oggetto dell'email

# Carica il corpo dell'email da un file HTM nella firma
signature_path = os.path.expandvars('%APPDATA%\\Microsoft\\Signatures\\')
signature_file = os.listdir(signature_path)[0]
signature_file_path = os.path.join(signature_path, signature_file)

with open(signature_file_path, 'r', encoding='utf-8') as file:
    signature_content = file.read()

# Aggiungi il contenuto della firma al corpo dell'email
mail.HTMLbody = f"Corpo dell'email: {signature_content}"



# Invia l'email
mail.Send()
