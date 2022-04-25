import sys, os
from PyQt5.QtCore import *
from PyQt5.QtWebEngineWidgets import *
from PyQt5.QtWidgets import QApplication, QAction
from win32com import client
from PyQt5 import QtPrintSupport, QtWidgets
import time
from typing import Callable

def donePrinting(boolean):
    print(boolean)

# os.startfile(url, 'print')

class PrintPrompt(QWebEngineView):
    def __init__(self):
        super(PrintPrompt, self).__init__()
        try:
            #self.web = QWebEngineView()
            url = r'C:\Users\simmsk\Desktop\templates\temp.html'
            self.load(QUrl.fromLocalFile(url))
        except Exception as e:
            print(e)

    def unshow(self, boolean):
        if boolean:
            print('Function called')
        try:
            self.close()
        except Exception as e:
            print(e)

if __name__=="__main__":
    # try:
    #     app = QApplication(sys.argv)
    #     web = QWebEngineView()
    #     url = r'C:\Users\simmsk\Desktop\templates\temp.html'
    #     web.load(QUrl.fromLocalFile(url))
    #     web.show()
    #     dialog = QtPrintSupport.QPrintDialog()
    #     if dialog.exec_() == QtWidgets.QDialog.Accepted:
    #         web.page().print(dialog.printer(), donePrinting)
    #         print('oof2')
    # except Exception as e:
    #     print(e)
    # finally:
    #     sys.exit(app.exec_())
    app = QApplication(sys.argv)
    web = PrintPrompt()
    web.show()
    web.setContextMenuPolicy(Qt.ActionsContextMenu)
    def printDialog():
        dialog = QtPrintSupport.QPrintDialog()
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            web.page().print(dialog.printer(), donePrinting)
            time.sleep(5)
    quitAction = QAction('Print', web)
    quitAction.triggered.connect(printDialog)
    web.addAction(quitAction)
    sys.exit(app.exec_())