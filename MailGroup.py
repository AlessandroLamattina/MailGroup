import os
import tempfile
from smb.SMBConnection import SMBConnection
from PyQt5.QtWidgets import (QStackedWidget,QDialog,QApplication,QMainWindow,QLabel,QLineEdit,QPushButton,QTextEdit,QMenuBar,QMenu,QAction,QStyleFactory,QVBoxLayout,QWidget,QHBoxLayout,QTableWidget,QTableWidgetItem,QLayout,QFileDialog,QFrame,QMessageBox)
from PyQt5.QtGui import QIcon,QPalette,QColor,QTextCharFormat
from PyQt5.QtGui import QTextCursor, QFont
from PyQt5.QtCore import Qt, QSize
import pyodbc
import win32com.client as win32
import validate_email

"""
1. cambio metodo della creazione delle tabelle tramite un pc muletto
"""
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MailSender")
        ##################################################################################################################################################
        self.conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\GroupMail.mdb;')
        self.cursor = self.conn.cursor()
        self.cursor.execute("SELECT numero, email, azienda, sesso, titolo, nome, cognome FROM Clienti")
        self.results = self.cursor.fetchall()
        self.email_cliente = {}
        self.aziende = []
        self.x = 1
        self.attachment_paths=[]
        for numero, email, azienda, sesso, titolo, nome, cognome in sorted(self.results, key=lambda y: y[0]):
            domain = azienda
            if domain not in self.email_cliente:
                self.aziende.append(azienda)
                self.email_cliente[domain] = []
            self.email_cliente[domain].append([numero,email,sesso,titolo,nome,cognome])
        self.email_deloitte = {}
        self.cursor2 = self.conn.cursor()
        self.cursor2.execute("SELECT azienda, persona1, persona2, persona3, persona4, persona5 FROM Dipendenti")
        self.results2 = self.cursor2.fetchall()
        for azienda_seguita, persona1, persona2, persona3, persona4, persona5 in self.results2:
            domain = azienda_seguita
            if domain not in self.email_deloitte:
                self.email_deloitte[domain] = []
            self.email_deloitte[domain].append([persona1, persona2, persona3, persona4, persona5])
        ##################################################################################################################################################
        self.current_path = os.getcwd()
        self.versione_software = 1.001
        # Ottieni il percorso della cartella superiore
        self.parent_path = os.path.abspath(os.path.join(self.current_path,os.pardir))
        # \\ITPC051911\mail_group
        self.server_name = "ITPC059100"
        self.server_ip = "ITPC059100"
        self.share_name = "mail_group"
        self.user_name = "Mail_Group_Updater"
        self.password = "Mail_Group_Updater"
        self.len_pre = 1
        self.len_pre2 = 1
        self.dizionario = []
        file_path = 'dizionario.txt'
        with open(file_path,'r',encoding='utf-8') as file:
            for parola in file:
                self.dizionario.append(parola.strip())
        self.initUI()
    def initUI(self):
        menubar = self.menuBar()
        fileMenu = menubar.addMenu("File")
        infoMenu = menubar.addMenu("Info")

        info_action = QAction("Info", self)
        info_action.triggered.connect(self.show_info)
        infoMenu.addAction(info_action)

        update_action = QAction("Check Update", self)
        update_action.triggered.connect(self.update)
        infoMenu.addAction(update_action)

        instr_action = QAction("Tabella Utenti", self)
        instr_action.triggered.connect(self.show_table)
        fileMenu.addAction(instr_action)

        mainwid_action = QAction("Schermata Principale",self)
        mainwid_action.triggered.connect(self.mainwidget)
        fileMenu.addAction(mainwid_action)

        send_action = QAction("Invia per mail", self)
        send_action.triggered.connect(self.send_emails)
        fileMenu.addAction(send_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        fileMenu.addAction(exit_action)
        self.mainwidget()
        self.popup= QMessageBox()
    def mainwidget(self):
        self.font = QFont()
        self.font.setPointSize(10)
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        ############################################################################
        self.subject_label = QLabel("Oggetto:",self)
        self.subject_label.setFont(self.font)
        self.layout.addWidget(self.subject_label)
        self.subject_entry = QLineEdit(self)
        self.subject_entry.setFont(self.font)
        self.layout.addWidget(self.subject_entry)
        ############################################################################
        self.body_label = QLabel("Testo:",self)
        self.body_label.setFont(self.font)
        self.layout.addWidget(self.body_label)
        self.body_entry = QTextEdit(self)
        self.body_entry.setStyleSheet("background-color: white;")
        self.body_entry.setFont(self.font)
        self.layout.addWidget(self.body_entry)
        #############################################################################
        self.file_button = QPushButton("Seleziona file",self)
        self.file_button.setFont(self.font)
        self.file_button.clicked.connect(self.select_files)
        self.layout.addWidget(self.file_button)
        self.file_button.setFixedWidth(200)
        self.files_layout = QHBoxLayout()
        self.files_layout.addWidget(self.file_button,Qt.AlignLeft)
        self.file_name_layout = QHBoxLayout()
        self.files_layout.addLayout(self.file_name_layout)
        self.files_layout.setAlignment(Qt.AlignLeft)
        self.layout.addLayout(self.files_layout)
        #############################################################################
        self.increment_button = QPushButton(">",self)
        self.increment_button.clicked.connect(self.increment_x)
        self.layout.addWidget(self.increment_button,Qt.AlignLeft)
        self.decrement_button = QPushButton("<",self)
        self.decrement_button.clicked.connect(self.decrement_x)
        self.layout.addWidget(self.decrement_button,Qt.AlignRight)
        self.decrement_button.setFixedWidth(100)
        self.increment_button.setFixedWidth(100)
        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(self.decrement_button)
        buttons_layout.addWidget(self.increment_button)
        buttons_layout.setSizeConstraint(QLayout.SetFixedSize)
        buttons_layout.setAlignment(Qt.AlignCenter)
        self.layout.addLayout(buttons_layout)
        #############################################################################
        self.preview_label = QLabel("Anteprima:",self)
        self.preview_label.setFont(self.font)
        self.layout.addWidget(self.preview_label)
        self.preview_text = QTextEdit(self)
        self.preview_text.setReadOnly(True)
        self.preview_text.setStyleSheet("background-color: white;")
        self.preview_text.setFont(self.font)
        self.layout.addWidget(self.preview_text)
        #############################################################################
        self.send_button = QPushButton("Invia",self)
        self.send_button.setFont(self.font)
        self.send_button.clicked.connect(self.send_emails)
        self.send_button.setFixedWidth(200)
        self.layout.addWidget(self.send_button)

        self.body_entry.textChanged.connect(self.update_preview)
        self.body_entry.textChanged.connect(self.verifica_parole)
    def verifica_parole(self):
        formato = QTextCharFormat()
        formato.setUnderlineColor(Qt.red)
        formato.setUnderlineStyle(QTextCharFormat.SpellCheckUnderline)
        body_template = self.body_entry.toPlainText()
        check = body_template.split(" ")
        format = QTextCharFormat()
        format.setUnderlineStyle(QTextCharFormat.SingleUnderline)
        format.setUnderlineColor(QColor("red"))
        if len(check) != self.len_pre:
            self.len_pre = len(check)
            for parola in check:
                if parola not in self.dizionario:
                    if "`" in parola \
                            or "!" in parola \
                            or "@" in parola \
                            or "#" in parola \
                            or "$" in parola \
                            or "%" in parola \
                            or "^" in parola \
                            or "&" in parola \
                            or "*" in parola \
                            or "(" in parola \
                            or ")" in parola \
                            or "-" in parola \
                            or "_" in parola \
                            or "=" in parola \
                            or "+" in parola \
                            or "[" in parola \
                            or "]" in parola \
                            or "{" in parola \
                            or "}" in parola \
                            or "|" in parola \
                            or "\\" in parola \
                            or ";" in parola \
                            or ":" in parola \
                            or "'" in parola \
                            or '"' in parola \
                            or "," in parola \
                            or "." in parola \
                            or "<" in parola \
                            or ">" in parola \
                            or "/" in parola \
                            or "?" in parola\
                            or parola.isdigit():
                        pass
                    else:
                        cursor = self.body_entry.document().find(parola)
                        cursor.mergeCharFormat(format)
                        cursor = self.body_entry.document().find(parola,cursor)
    def increment_x(self):
        print(self.x)
        if self.x < len(self.aziende):
            self.x += 1
            self.update_preview()
    def decrement_x(self):
        if self.x > 1:
            self.x -= 1
            self.update_preview()
    def crate_table(self):
        self.button_table_widget = QWidget()
        self.setCentralWidget(self.button_table_widget)
        self.button_table_layout = QVBoxLayout(self.button_table_widget)
    def show_table(self):
        self.button_table_widget = QWidget()
        self.setCentralWidget(self.button_table_widget)
        self.button_table_layout= QVBoxLayout(self.button_table_widget)
        azienda_layout = QHBoxLayout()
        for azienda in self.aziende:
            self.azienda_button = QPushButton(f"{azienda}",self)
            self.azienda_button.clicked.connect(lambda: self.open_tab())
            self.azienda_button.setProperty("azienda", azienda)
            azienda_layout.addWidget(self.azienda_button)
            self.button_table_layout.addLayout(azienda_layout)
        back_button = QPushButton("Schermata Principale",self)
        back_button.clicked.connect(self.mainwidget)
        back_button.setFixedWidth(150)
        self.button_table_layout.addWidget(back_button)
    def open_tab(self):
        azienda = self.sender().property("azienda")
        self.second_widget = QWidget()
        self.setCentralWidget(self.second_widget)
        self.layout2 = QVBoxLayout(self.second_widget)
        self.table = QTableWidget(self)
        self.table.setRowCount(len(self.email_cliente[azienda]))
        self.table.setColumnCount(5)
        for i in range(0,len(self.email_cliente[azienda])):
            for j in range(1,6):
                self.table.setItem(i,j - 1,QTableWidgetItem(self.email_cliente[azienda][i][j]))
        self.layout2.addWidget(self.table)
        back_button = QPushButton("Schermata Principale",self)
        back_button.clicked.connect(self.mainwidget)
        back_button.setFixedWidth(150)
        self.layout2.addWidget(back_button)
    def update_preview(self):
        body_template = self.body_entry.toPlainText()
        body = body_template.replace("{email}",self.email_cliente[self.aziende[self.x - 1]][0][1]) \
            .replace("{sesso}",self.email_cliente[self.aziende[self.x - 1]][0][2]) \
            .replace("{titolo}",self.email_cliente[self.aziende[self.x - 1]][0][3]) \
            .replace("{nome}",self.email_cliente[self.aziende[self.x - 1]][0][4]) \
            .replace("{cognome}",self.email_cliente[self.aziende[self.x - 1]][0][5]) \
            .replace("{azienda}",self.aziende[self.x - 1])
        self.preview_text.setPlainText(body)

    def select_files(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_dialog = QFileDialog()
        file_paths,_ = file_dialog.getOpenFileNames(self,"Seleziona Allegati","","All Files (*)",options=options)
        if file_paths:
            for file_path in file_paths:
                self.attachment_paths.append(file_path)
            file_names = self.attachment_paths
            for i in range(0,len(file_names)):
                file_name = file_names[i]
                file_name2 = file_names[i].split('/')[-1]
                self.file_name_button = QPushButton(f"{file_name2}",self)
                self.file_name_button.clicked.connect(lambda: self.delete_file())
                self.file_name_button.setFont(self.font)
                self.file_name_button.setFixedWidth(150)
                self.file_name_button.setProperty("file_name",file_name)
                self.file_name_layout.addWidget(self.file_name_button)
    def delete_file(self):
        file = self.sender().property("file_name")
        button_to_remove = self.sender()
        self.file_name_layout.removeWidget(button_to_remove)
        self.attachment_paths.remove(file)
        button_to_remove.deleteLater()
    def send_emails(self):
        send_error=[]
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        subject = self.subject_entry.text()
        body_template = self.body_entry.toPlainText()
        for domain, recipients in self.email_cliente.items():
            message = outlook.CreateItem(0)
            message.Subject = subject
            recipients_list = message.Recipients
            for domain_f, recipients_f in self.email_deloitte.items():
                i = 0
                jj = []
                if domain_f == domain:
                    for j in recipients_f:
                        for ii in j:
                            if ii is not None:
                                jj.append(ii)
                    else:
                        pass
                    for numero,email,sesso,titolo,nome,cognome in recipients:
                        if validate_email.validate_email(email):
                            recipient_email = namespace.CreateRecipient(email)
                            if i == 0:
                                recipients_list.Add(recipient_email).Type = 1
                            else:
                                recipients_list.Add(recipient_email).Type = 2
                            body = body_template.replace("{sesso}", sesso)\
                                                .replace("{nome}", nome)\
                                                .replace("{cognome}", cognome)\
                                                .replace("{email}", email)\
                                                .replace("{titolo}", titolo)\
                                                .replace("{azienda}", domain)
                            if i == 0:
                                message.Body = body
                                i += 1
                            signature_path = os.path.expandvars('%APPDATA%\\Microsoft\\Signatures\\')
                            signature_file = os.listdir(signature_path)[0]
                            with open(signature_path + signature_file,'r') as file:
                                signature = file.read()

                            message.HTMLBody += f"<br><img src='{signature}'width='120' height='30'>"
                    for email in jj:
                        recipient_email = namespace.CreateRecipient(email)
                        recipients_list.Add(recipient_email).Type = 2
                    recipients_list.ResolveAll()
                    try:
                        if self.attachment_paths:
                            for attachment in self.attachment_paths:
                                message.Attachments.Add(attachment)

                        message.Send()
                        send_result = 0
                        print("Email inviata con successo a", domain)
                    except Exception as e:
                        print("Errore durante l'invio dell'email a", domain)
                        send_error.append(domain)
                        print(str(e))
                        send_result = 1
            if send_result == 0:
                self.popup.setWindowTitle("Invio mail")
                self.popup.setText("Email inviate con successo.")
                self.popup.setIcon(QMessageBox.Information)
                #self.popup.addButton("OK",QMessageBox.AcceptRole)
                self.popup.exec_()
            elif send_result == 1:
                self.popup.setWindowTitle("Invio non riuscito")
                self.popup.setText("Ci sono stati problemi nell'inviare la mail a:", send_error)
                self.popup.setIcon(QMessageBox.Warning)
                #self.popup.addButton("OK",QMessageBox.AcceptRole)
                self.popup.exec_()
        self.conn.close()
    def show_info(self):
        self.popup.setWindowTitle("Info")
        self.popup.setText(f"Questo software è stato creato da Alessandro Lamattina.\nLa versione attuale è la {self.versione_software}."
                           f"\nIn caso di richieste di nuove funzioni \no di segnalazioni bug scrivere alla mail alamattinalarocca@deloitte.it")
        self.popup.setIcon(QMessageBox.Information)
        self.popup.exec_()
    def update(self):
        try:
            # Effettua la connessione
            conn = SMBConnection(self.user_name,self.password,"localhost",self.server_name,use_ntlm_v2=True)
            conn.connect(self.server_ip)
            # Scarica il file e salvalo in una directory temporanea
            try:
                file_obj = tempfile.NamedTemporaryFile(delete=False)
                file_path = file_obj.name
                file_obj.close()
                with open(file_path,'wb') as fp:
                    conn.retrieveFile(self.share_name,"versione_file.txt",fp)
                with open(file_path,'r') as fp:
                    versione_server = fp.readlines()
                    if float(versione_server[0]) == self.versione_software:
                        self.popup.setWindowTitle("Versione")
                        self.popup.setText("La tua versione è la più aggiornata al momento.")
                        self.popup.setIcon(QMessageBox.Information)
                        self.popup.exec_()
                    elif float(versione_server[0]) > self.versione_software:
                        os.startfile("update\\updater.exe")
                        self.popup.setWindowTitle("Versione")
                        self.popup.setText("Aggiornamento in corso.")
                        self.popup.setIcon(QMessageBox.Information)
                        self.popup.exec_()
            finally:
                os.unlink(file_path)
            # Chiudi la connessione
            conn.close()
        except:
            self.popup.setWindowTitle("Errore")
            self.popup.setText("Problemi a raggiungere il server")
            self.popup.setIcon(QMessageBox.Warning)
            self.popup.exec_()
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))  # Applica uno stile moderno
    window = MainWindow()
    window.show()
    icon = QIcon(".\icona.png")  # Sostituisci  con il percorso dell'icona desiderata
    window.setWindowIcon(icon)
    window.resize(1920,1055)
    sys.exit(app.exec())
