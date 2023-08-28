import pyodbc
from tkinter import Tk, Label, Entry, Button, Text, Menu, ttk
import validate_email
import win32com.client as win32


class EmailSender:
    def __init__(self):

        ##################################################################################################################################################
        self.conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\test.mdb;')
        self.cursor = self.conn.cursor()
        self.cursor.execute("SELECT numero, email, azienda, sesso, titolo, cognome FROM TestEmail")
        self.results = self.cursor.fetchall()
        self.email_dict = {}
        self.aziende = []
        self.x = 1

        for numero, email, azienda, sesso, titolo, cognome in sorted(self.results, key=lambda y: y[0]):
            domain = azienda
            if domain not in self.email_dict:
                self.aziende.append(azienda)
                self.email_dict[domain] = []
            self.email_dict[domain].append([numero, email, sesso, titolo, cognome])
        #print(len(self.aziende))
        ##################################################################################################################################################
        self.initUI()


    def initUI(self):
        """self.root = Tk()
        self.root.title("Invio Offerte")
        self.root.geometry("500x600")"""

        self.root = Tk()
        self.root.title("Invio Offerte")


        style = ttk.Style()
        style.configure("TFrame",background="green")  # Imposta il colore di sfondo su verde


        menubar = Menu(self.root)
        self.root.config(menu=menubar)



        fileMenu = Menu(menubar)
        infoMenu = Menu(menubar)
        fileMenu.add_command(label="Info",command=self.contatti)
        infoMenu.add_command(label="Info",command=self.info)
        infoMenu.add_command(label="Istruzioni",command=self.istruzioni)
        fileMenu.add_command(label="Invia per mail", command=self.mail)
        fileMenu.add_command(label="Exit",command=self.onExit)
        menubar.add_cascade(label="File", menu=fileMenu)
        menubar.add_cascade(label="Info",menu=infoMenu)
        style = ttk.Style()
        style.theme_use('clam')

        self.subject_label = Label(self.root,text="Oggetto:")
        self.subject_label.pack()
        self.subject_entry = Entry(self.root)
        self.subject_entry.pack()

        self.body_label = Label(self.root,text="Testo:")
        self.body_label.pack()
        self.body_entry = Text(self.root,height=10,width=40)
        self.body_entry.pack()

        self.increment_button = Button(self.root,text=">",command=self.increment_x)
        self.increment_button.pack()

        self.decrement_button = Button(self.root,text="<",command=self.decrement_x)
        self.decrement_button.pack()

        self.preview_label = Label(self.root,text="Anteprima:")
        self.preview_label.pack()
        self.preview_text = Text(self.root,height=10,width=40)
        self.preview_text.pack()

        self.send_button = Button(self.root,text="Invia",command=self.send_emails)
        self.send_button.pack()


        self.body_entry.bind("<KeyRelease>",self.update_preview)
    def increment_x(self):
        print(self.x)
        if self.x < len(self.aziende):
            self.x += 1
            self.update_preview(None)
    def decrement_x(self):
        if self.x > 1:
            self.x -= 1
            self.update_preview(None)
    def update_preview(self, event):

        body_template = self.body_entry.get("1.0", "end-1c")
        body = body_template\
        .replace("{sesso}", self.email_dict[self.aziende[self.x-1]][0][2])\
        .replace("{cognome}", self.email_dict[self.aziende[self.x-1]][0][4])\
        .replace("{email}", self.email_dict[self.aziende[self.x-1]][0][1])\
        .replace("{titolo}", self.email_dict[self.aziende[self.x-1]][0][3])\
        .replace("{azienda}", self.aziende[self.x-1])

        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", body)

    def send_emails(self):
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")

        subject = self.subject_entry.get()
        body_template = self.body_entry.get("1.0","end-1c")

        email_f = {}

        cursor2 = self.conn.cursor()
        cursor2.execute("SELECT azienda, persona1, persona2, persona3, persona4, persona5 FROM Tabella2")
        results2 = cursor2.fetchall()

        for azienda_seguita,persona1,persona2,persona3,persona4,persona5 in results2:
            domain = azienda_seguita
            if domain not in self.email_dict:
                email_f[domain] = []
            email_f[domain].append([persona1,persona2,persona3,persona4,persona5])

        for domain,recipients in self.email_dict.items():
            message = outlook.CreateItem(0)
            message.Subject = subject

            recipients_list = message.Recipients
            for domain_f,recipients_f in email_f.items():
                i = 0
                jj = []
                if domain_f == domain:
                    for j in recipients_f:
                        for ii in j:
                            if ii is not None:
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

                    # Adding attachments
                    attachment_paths = self.get_attachment_paths()  # Get the attachment file paths from the user
                    for attachment_path in attachment_paths:
                        attachment = message.Attachments.Add(Source=attachment_path)

                    try:
                        message.Send()
                        print("Email inviata con successo a",domain)
                    except Exception as e:
                        print("Errore durante l'invio dell'email a",domain)
                        print(str(e))

        self.conn.close()

    def get_attachment_paths(self):
        # Open a file dialog to select one or more attachment files
        attachment_paths = self.filedialog.askopenfilenames()
        return attachment_paths

    def run(self):
        self.root.mainloop()
    def onExit(self):
        self.quit()

    def restart_screen(self):
        self.master.update()  # Distruggi la finestra principale

    def istruzioni(self):
        i = 0

    def mail(self):
        i = 0

    def contatti(self):
        i = 0

    def info(self):
        i = 0
if __name__ == "__main__":
    email_sender = EmailSender()
    email_sender.run()

