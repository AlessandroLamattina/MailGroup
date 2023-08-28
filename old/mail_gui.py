import pyodbc
import win32com.client as win32
from tkinter import Tk, Label, Entry, Button, Text
import validate_email

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\test.mdb;')
cursor = conn.cursor()

cursor.execute("SELECT numero, email, azienda, sesso, titolo, cognome FROM TestEmail")
results = cursor.fetchall()


def send_emails():
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\test.mdb;')
    cursor = conn.cursor()

    cursor.execute("SELECT numero, email, azienda, sesso, titolo, cognome FROM TestEmail")  # TestEmail
    results = cursor.fetchall()

    cursor.execute("SELECT azienda, persona1,persona2,persona3,persona4,persona5 FROM Tabella2")
    results2 = cursor.fetchall()

    outlook = win32.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")

    subject = subject_entry.get()
    body_template = body_entry.get("1.0","end-1c")

    email_dict = {}
    email_f={}

    for azienda_seguita, persona1,persona2,persona3,persona4,persona5  in results2:
        domain = azienda_seguita
        if domain not in email_dict:
            email_f[domain] = []
        email_f[domain].append([persona1,persona2,persona3,persona4,persona5])

    for numero,email,azienda,sesso,titolo,cognome in sorted(results,key=lambda x: x[0]):
        domain = azienda
        if domain not in email_dict:
            email_dict[domain] = []
        email_dict[domain].append([numero,email,sesso,titolo,cognome])

    for domain,recipients in email_dict.items():
        message = outlook.CreateItem(0)
        message.Subject = subject

        recipients_list = message.Recipients
        for domain_f,recipients_f in email_f.items():
            i = 0
            jj = []
            if domain_f == domain:
                for j in recipients_f:
                    for ii in j:
                        if ii != None:
                            jj.append(ii)
                else:
                    print("Indirizzo email non valido:",email)
                for numero,email,sesso,titolo,cognome in recipients:
                    if validate_email.validate_email(email):
                        recipient_email = namespace.CreateRecipient(email)
                        if i == 0:
                            recipients_list.Add(recipient_email).Type = 1
                        else:
                            recipients_list.Add(recipient_email).Type = 2
                        body = body_template.replace("{sesso}",sesso).replace("{cognome}",cognome).replace(
                            "{email}",email).replace("{titolo}",titolo).replace("{azienda}",domain)
                        if i == 0:
                            message.Body = body
                            i += 1
                for email in jj:
                    recipient_email = namespace.CreateRecipient(email)
                    recipients_list.Add(recipient_email).Type = 2
                recipients_list.ResolveAll()
                """try:
                    message.Send()
                    print("Email inviata con successo a",domain)
                except Exception as e:
                    print("Errore durante l'invio dell'email a",domain)
                    print(str(e))
                conn.close()
"""
root = Tk()
root.title("Invio Offerte")
root.geometry("500x600")

# Etichette e campi di input per oggetto e corpo del messaggio
subject_label = Label(root,text="Oggetto:")
subject_label.pack()
subject_entry = Entry(root)
subject_entry.pack()

body_label = Label(root,text="Testo:")
body_label.pack()
body_entry = Text(root,height=10,width=40)
body_entry.pack()

x = 0

def increment_x():
    global x
    x += 1
    update_preview(None)  # Passa un valore fittizio per l'argomento event

def decrement_x():
    global x
    x -= 1
    update_preview(None)  # Passa un valore fittizio per l'argomento event

increment_button = Button(root, text=">", command=increment_x)
increment_button.pack()

decrement_button = Button(root, text="<", command=decrement_x)
decrement_button.pack()

# Area di anteprima
preview_label = Label(root,text="Anteprima:")
preview_label.pack()
preview_text = Text(root,height=10,width=40)
preview_text.pack()

def update_preview(event):
    body_template = body_entry.get("1.0","end-1c")
    email_dict = {}
    global x

    for numero,email,azienda,sesso,titolo,cognome in sorted(results,key=lambda y: y[0]):
        domain = azienda
        if domain not in email_dict:
            email_dict[domain] = []
        email_dict[domain].append([numero,email,sesso,titolo,cognome])
    print(email_dict.get(x, [[""]])[0][4])
    body = body_template\
            .replace("{sesso}", email_dict[x][0][2])\
            .replace("{cognome}",email_dict[x][0][4])\
            .replace("{email}",email[x])\
            .replace("{titolo}",titolo[x])\
            .replace("{azienda}",domain[x])

    preview_text.delete("1.0","end")
    preview_text.insert("1.0",body)

body_entry.bind("<KeyRelease>",update_preview)

# Pulsante di invio
send_button = Button(root,text="Invia",command=send_emails)
send_button.pack()

root.mainloop()


"""
import pyodbc
import win32com.client as win32
from tkinter import Tk, Label, Entry, Button, Text
import validate_email

def send_emails():

    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\test.mdb;')
    cursor = conn.cursor()

    cursor.execute("SELECT numero, email, azienda, sesso, titolo, cognome FROM TestEmail")  # TestEmail
    results = cursor.fetchall()

    cursor.execute("SELECT azienda, persona1,persona2,persona3,persona4,persona5 FROM Tabella2")
    results2 = cursor.fetchall()

    outlook = win32.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")

    subject = subject_entry.get()
    body_template = body_entry.get("1.0","end-1c")

    email_dict = {}
    email_f={}

    for azienda_seguita, persona1,persona2,persona3,persona4,persona5  in results2:
        domain = azienda_seguita
        if domain not in email_dict:
            email_f[domain] = []
        email_f[domain].append([persona1,persona2,persona3,persona4,persona5])

    for numero,email,azienda,sesso,titolo,cognome in sorted(results,key=lambda x: x[0]):
        domain = azienda
        if domain not in email_dict:
            email_dict[domain] = []
        email_dict[domain].append([numero,email,sesso,titolo,cognome])

    for domain,recipients in email_dict.items():
        message = outlook.CreateItem(0)
        message.Subject = subject

        recipients_list = message.Recipients
        for domain_f,recipients_f in email_f.items():
            i = 0
            jj = []
            if domain_f == domain:
                for j in recipients_f:
                   for ii in j:
                       if ii != None:
                           jj.append(ii)
                else:
                    print("Indirizzo email non valido:",email)
                for numero,email,sesso,titolo,cognome in recipients:
                    if validate_email.validate_email(email):
                        recipient_email = namespace.CreateRecipient(email)
                        if i == 0:
                            recipients_list.Add(recipient_email).Type = 1
                        else:
                            recipients_list.Add(recipient_email).Type = 2
                        body = body_template.replace("{sesso}",sesso).replace("{cognome}",cognome).replace("{email}",email).replace("{titolo}",titolo).replace("{azienda}",domain)
                        if i == 0:
                            message.Body = body
                            i += 1
                for email in jj:
                    recipient_email = namespace.CreateRecipient(email)
                    recipients_list.Add(recipient_email).Type = 2
                recipients_list.ResolveAll()
                try:
                    message.Send()
                    print("Email inviata con successo a",domain)
                except Exception as e:
                    print("Errore durante l'invio dell'email a",domain)
                    print(str(e))
    conn.close()

root = Tk()
root.title("Invio Offerte")
root.geometry("400x300")

# Etichette e campi di input per oggetto e corpo del messaggio
subject_label = Label(root, text="Oggetto:")
subject_label.pack()
subject_entry = Entry(root)
subject_entry.pack()

body_label = Label(root, text="Testo:")
body_label.pack()
body_entry = Text(root, height=10, width=40)
body_entry.pack()

# Pulsante di invio
send_button = Button(root, text="Invia", command=send_emails)
send_button.pack()

root.mainloop()
"""