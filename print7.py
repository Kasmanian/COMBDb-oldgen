import sys, os
from PyQt5.QtCore import *
from PyQt5.QtWebEngineWidgets import *
from PyQt5.QtWidgets import QApplication, QAction
from win32com import client
from PyQt5 import QtPrintSupport, QtWidgets
import time
from typing import Callable

def donePrinting(boolean):
    print('oof')

class PrintPrompt(QApplication):
    def __init__(self, argv):
        super(PrintPrompt, self).__init__(argv)
        self.web = QWebEngineView()
        self.web.setContextMenuPolicy(Qt.ActionsContextMenu)
        quitAction = QAction('Print', self.web)
        quitAction.triggered.connect(self.printDialogue)
        self.web.addAction(quitAction)
        url = r'C:\Users\simmsk\Desktop\templates\temp.html'
        self.web.setWindowTitle(url)
        self.web.load(QUrl.fromLocalFile(url))
        self.web.show()
        sys.exit(self.exec_())

    def printDialogue(self):
        self.dialog = QtPrintSupport.QPrintDialog()
        if self.dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.web.page().print(self.dialog.printer(), donePrinting)
            print('oof2')

if __name__=="__main__":
    printPrompt = PrintPrompt(sys.argv)