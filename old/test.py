from PyQt5.QtGui import QTextCharFormat,QTextCursor,QBrush,QColor,QFont
from PyQt5.QtWidgets import QApplication,QMainWindow,QTextEdit
from PyQt5.QtCore import Qt,QEvent


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.len_pre=1
        self.dizionario = []
        file_path = 'dizionario.txt'
        with open(file_path,'r',encoding='utf-8') as file:
            for parola in file:
                self.dizionario.append(parola.strip())
#################################################################################################
        self.text_edit = QTextEdit()
        self.setCentralWidget(self.text_edit)
        #self.text_edit.installEventFilter(self)
        self.text_edit.textChanged.connect(self.verifica_parole)
    def verifica_parole(self):
        formato = QTextCharFormat()
        formato.setUnderlineColor(Qt.red)
        formato.setUnderlineStyle(QTextCharFormat.SpellCheckUnderline)
        body_template = self.text_edit.toPlainText()
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
                        cursor = self.text_edit.document().find(parola)
                        cursor.mergeCharFormat(format)
                        cursor = self.text_edit.document().find(parola,cursor)

if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
