import sys, os
from PyQt5.QtCore import *
from PyQt5.QtWebEngineWidgets import *
from PyQt5.QtWidgets import *
from win32com import client
from PyQt5 import QtGui
import time

class MyBrowser(QWebEnginePage):

    def userAgentForUrl(self, url):
        return "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36"

class Web(QWebEngineView):

    def load(self, url):
        self.setUrl(QUrl(url))

    def adjustTitle(self):
        self.setWindowTitle(self.title())

    def disableJS(self):
        settings = QWebEngineSettings.globalSettings()
        settings.setAttribute(QWebEngineSettings.JavascriptEnabled, False)

class Main(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Name')
        # self.setWindowIcon(QtGui.QIcon('icon.png'))

        self.btn = QPushButton('Button', self)
        self.btn.resize(self.btn.sizeHint())
        self.btn.move(20, 20)
        self.show()

app = QApplication(sys.argv)
web = Web()
main = Main()
url = r'C:\Users\simmsk\Desktop\templates\temp.html'
web.load(QUrl.fromLocalFile(url))
web.show()
sys.exit(app.exec_())