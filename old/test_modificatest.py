import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLineEdit, QPushButton, QVBoxLayout, QWidget

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Esempio TextBox PyQt")
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()

        self.textbox = QLineEdit(self)
        layout.addWidget(self.textbox)

        self.button = QPushButton("Modifica Testo", self)
        self.button.clicked.connect(self.modify_text)
        layout.addWidget(self.button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def modify_text(self):
        new_text = "Nuovo testo modificato"
        self.textbox.setText(new_text)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

