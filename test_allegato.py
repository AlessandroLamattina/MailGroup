from PyQt5.QtWidgets import (
    QStackedWidget,QDialog,QApplication,QMainWindow,QLabel,QLineEdit,
    QPushButton,QTextEdit,QMenuBar,QMenu,QAction,QStyleFactory,
    QVBoxLayout,QWidget,QHBoxLayout,QTableWidget,QTableWidgetItem,QLayout,QFileDialog
)
from PyQt5.QtGui import QIcon
from PyQt5.QtGui import QTextCursor, QFont
from PyQt5.QtCore import Qt
import pyodbc
import win32com.client as win32
import validate_email

"""
1. aggiungere invio allegati
2. cambio metodo della creazione delle tabelle tramite un pc muletto
3. modifica delle tabelle da software stesso (quindi cambio metodo del richiamo dei dati)

"""
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Invio Offerte")
        ##################################################################################################################################################
        self.conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\test.mdb;')
        self.cursor = self.conn.cursor()
        self.cursor.execute("SELECT numero, email, azienda, sesso, titolo, nome, cognome FROM TestEmail")
        self.results = self.cursor.fetchall()
        self.email_dict = {}
        self.aziende = []
        self.x = 1

        for numero, email, azienda, sesso, titolo, nome, cognome in sorted(self.results, key=lambda y: y[0]):
            domain = azienda
            if domain not in self.email_dict:
                self.aziende.append(azienda)
                self.email_dict[domain] = []
            self.email_dict[domain].append([numero, email, sesso, titolo, nome, cognome])
        ##################################################################################################################################################
        self.initUI()
    def initUI(self):

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu("File")
        infoMenu = menubar.addMenu("Info")

        info_action = QAction("Info", self)
        info_action.triggered.connect(self.show_info)
        infoMenu.addAction(info_action)

        instr_action = QAction("Tabella Utenti", self)
        instr_action.triggered.connect(self.show_table)
        infoMenu.addAction(instr_action)

        send_action = QAction("Invia per mail", self)
        send_action.triggered.connect(self.send_emails)
        fileMenu.addAction(send_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        fileMenu.addAction(exit_action)

        self.subject_label = QLabel("Oggetto:",self)
        self.layout.addWidget(self.subject_label)
        self.subject_entry = QLineEdit(self)
        self.layout.addWidget(self.subject_entry)

        self.body_label = QLabel("Testo:",self)
        self.layout.addWidget(self.body_label)
        self.body_entry = QTextEdit(self)
        self.body_entry.setStyleSheet("background-color: white;")
        self.layout.addWidget(self.body_entry)

        self.file_button = QPushButton("Seleziona file",self)
        self.file_button.clicked.connect(self.select_files)
        self.label_files = QLabel("", self)
        files_layout = QHBoxLayout()  #Layout orizzontale per i bottoni
        files_layout.addWidget(self.file_button)
        files_layout.addWidget(self.label_files)
        self.layout.addLayout(files_layout)

        self.increment_button = QPushButton(">",self)
        self.increment_button.clicked.connect(self.increment_x)
        self.layout.addWidget(self.increment_button, Qt.AlignLeft)

        self.decrement_button = QPushButton("<",self)
        self.decrement_button.clicked.connect(self.decrement_x)
        self.layout.addWidget(self.decrement_button, Qt.AlignRight)
        self.decrement_button.setFixedWidth(100)
        self.increment_button.setFixedWidth(100)
        buttons_layout = QHBoxLayout()  # Layout orizzontale per i bottoni
        buttons_layout.addWidget(self.decrement_button)
        buttons_layout.addWidget(self.increment_button)
        buttons_layout.setSizeConstraint(QLayout.SetFixedSize)  # Imposta la grandezza fissa per il layout
        buttons_layout.setAlignment(Qt.AlignCenter)
        self.layout.addLayout(buttons_layout)

        self.preview_label = QLabel("Anteprima:",self)
        self.layout.addWidget(self.preview_label)
        self.preview_text = QTextEdit(self)
        self.preview_text.setReadOnly(True)
        self.preview_text.setStyleSheet("background-color: white;")
        self.layout.addWidget(self.preview_text)

        self.send_button = QPushButton("Invia",self)
        self.send_button.clicked.connect(self.send_emails)
        self.layout.addWidget(self.send_button)
        #self.second_window_button = QPushButton("Apri Seconda Schermata",self)
        #self.second_window_button.clicked.connect(self.showSecondWindows)
        #self.layout.addWidget(self.second_window_button)

        self.body_entry.textChanged.connect(self.update_preview)

        # Creazione dei bottoni
    def increment_x(self):
        print(self.x)
        if self.x < len(self.aziende):
            self.x += 1
            self.update_preview()
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
        back_button.clicked.connect(self.initUI)
        back_button.setFixedWidth(150)
        self.button_table_layout.addWidget(back_button)
    def decrement_x(self):
        if self.x > 1:
            self.x -= 1
            self.update_preview()

    def update_preview(self):
        body_template = self.body_entry.toPlainText()
        body = body_template.replace("{sesso}", self.email_dict[self.aziende[self.x - 1]][0][2]) \
            .replace("{cognome}", self.email_dict[self.aziende[self.x - 1]][0][4]) \
            .replace("{email}", self.email_dict[self.aziende[self.x - 1]][0][1]) \
            .replace("{titolo}", self.email_dict[self.aziende[self.x - 1]][0][3]) \
            .replace("{azienda}", self.aziende[self.x - 1])
        self.preview_text.setPlainText(body)

    def select_files(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_dialog = QFileDialog()
        file_paths,_ = file_dialog.getOpenFileNames(self,"Seleziona Allegati","","All Files (*)",options=options)
        if file_paths:
            self.attachment_paths = file_paths
            print(file_paths)
            #self.label_files.setText(file_paths[0].split("/")[-1])
            i = 0
            for file_path in file_paths:
                file_name = f"{file_name}, {file_path[i].split('/')[-1]}"
                i += 1
            #file_path = string(file_paths)
            #print(file_paths[0].split("/")[-1])
    def send_emails(self):
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")

        subject = self.subject_entry.text()
        body_template = self.body_entry.toPlainText()

        email_f = {}

        cursor2 = self.conn.cursor()
        cursor2.execute("SELECT azienda, persona1, persona2, persona3, persona4, persona5 FROM Tabella2")
        results2 = cursor2.fetchall()

        for azienda_seguita, persona1, persona2, persona3, persona4, persona5 in results2:
            domain = azienda_seguita
            if domain not in self.email_dict:
                email_f[domain] = []
            email_f[domain].append([persona1, persona2, persona3, persona4, persona5])

        for domain, recipients in self.email_dict.items():
            message = outlook.CreateItem(0)
            message.Subject = subject

            recipients_list = message.Recipients
            for domain_f, recipients_f in email_f.items():
                i = 0
                jj = []
                if domain_f == domain:
                    for j in recipients_f:
                        for ii in j:
                            if ii is not None:
                                jj.append(ii)
                    else:
                        print("Indirizzo email non valido:", email)
                    for numero, email, sesso, titolo, cognome in recipients:
                        if validate_email.validate_email(email):
                            recipient_email = namespace.CreateRecipient(email)
                            if i == 0:
                                recipients_list.Add(recipient_email).Type = 1
                            else:
                                recipients_list.Add(recipient_email).Type = 2
                            body = body_template.replace("{sesso}", sesso).replace("{cognome}", cognome).replace(
                                "{email}", email).replace("{titolo}", titolo).replace("{azienda}", domain)
                            if i == 0:
                                message.Body = body
                                i += 1
                    for email in jj:
                        recipient_email = namespace.CreateRecipient(email)
                        recipients_list.Add(recipient_email).Type = 2
                    recipients_list.ResolveAll()

                    try:

                        if self.attachment_path:
                            for attachment in self.attachment_path:
                                message.Attachments.Add(attachment)
                        message.Send()
                        print("Email inviata con successo a", domain)
                    except Exception as e:
                        print("Errore durante l'invio dell'email a", domain)
                        print(str(e))
                    self.conn.close()

        self.conn.close()

    def show_info(self):
        # Resto del codice per visualizzare le informazioni...
        pass
    def open_tab(self):
        azienda = self.sender().property("azienda")
        print(azienda)
        self.second_widget = QWidget()
        self.setCentralWidget(self.second_widget)
        self.layout2 = QVBoxLayout(self.second_widget)
        self.table = QTableWidget(self)
        self.table.setRowCount(len(self.email_dict[azienda]))
        self.table.setColumnCount(5)
        """        header_labels = ["E-Mail","Saluto","Titolo","Nome","Cognome"]
                for label in header_labels:
                    self.table.setHorizontalHeaderLabels(label)
        """
        for i in range(0,len(self.email_dict[azienda])):
            for j in range(1,6):
                self.table.setItem(i, j-1, QTableWidgetItem(self.email_dict[azienda][i][j]))
        self.layout2.addWidget(self.table)
        back_button = QPushButton("Schermata Principale",self)
        back_button.clicked.connect(self.initUI)
        back_button.setFixedWidth(150)
        self.layout2.addWidget(back_button)

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
