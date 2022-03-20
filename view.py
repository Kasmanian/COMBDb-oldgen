from tkinter import CENTER, Button
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
import sys

class View(QApplication):
    def __init__(self, model):
        app = QApplication(sys.argv)
        welcome = Window(model)
        widget = QtWidgets.QStackedWidget()
        widget.addWidget(welcome)
        widget.setFixedHeight(1200)
        widget.setFixedWidth(1600)
        widget.show()
        try:
            sys.exit(app.exec())
        except:
            print("Exiting")


class Window(QMainWindow):
    def __init__(self, model):
        super(Window, self).__init__()
        self.model = model
        loadUi("COMBdb/UI Screens/COMBdb_Login.ui", self)
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.login.clicked.connect(self.handleLoginPressed)

    def handleLoginPressed(self):
        if self.model.adminLogin(self.usrnm.text(), self.pswd.text()):
            print('Success! Logging you in...')
            return
        print('Wrong username or password')