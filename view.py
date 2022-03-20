from tkinter import CENTER, Button
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
import sys

class View(QApplication):
    def __init__(self):
        app = QApplication(sys.argv)
        welcome = Window()
        widget = QtWidgets.QStackedWidget()
        widget.addWidget(welcome)
        widget.setFixedHeight(1200)
        widget.setFixedWidth(1600)
        widget.show()
        try:
            sys.exit(app.exec())
        except:
            print("Exiting")
        #super(View, self).__init__(sys.argv)
        #self.win = Window()
        #Window()

    #def init(self):
        #sys.exit(self.exec_())


class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        loadUi("COMBdb/UI Screens/COMBdb_Login.ui", self)
        #self.initUI()

    #def initUI(self):
        #self.setWindowTitle('COMBdb')
        #self.showMaximized()

      
      

      
      

