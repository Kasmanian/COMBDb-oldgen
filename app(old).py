from PyQt5.QtWidgets import QApplication, QStackedWidget
from QtClasses.QForms import *
from QtClasses.QNavs import *
from QtClasses.QPopups import *
from datetime import date
import sys

def passPrintPrompt(boolean):
        pass

class App:
    def __init__(self, db):
        self.db = db
        app = QApplication(sys.argv)
        app.setApplicationDisplayName('COMBDb')
        screen = QLogin(self)
        self.widget = QStackedWidget()
        self.widget.addWidget(screen)
        self.widget.setGeometry(10,10,1000,800)
        self.widget.showMaximized()
        if not self.db.connect():
            self.showFileBrowser('DBQ', 'MS Access Files (*.accdb)')
        else:
            pass
            # self.setClinicianList()
            self.date = date.today()
        try:
            sys.exit(app.exec())
        except Exception as e:
            self.showErrorScreen(e)

    def showFileBrowserPopup(self, key, ext):
        self.QFileBrowser = QFileBrowser(self, key, ext)
        self.QFileBrowser.show()

    def showErrorPopup(self, message):
        self.QError = QError(self, message)
        self.QError.show()

    def showConfirmationPopup(self):
        self.QConfirmation = QConfirmation(self)
        self.QConfirmation.show()

    def showArchiverPopup(self):
        self.QArchiver = QArchiver(self)
        self.QArchiver.show()

    def showLoginForm(self):
        QLogin = QLogin(self)
        self.widget.addWidget(QLogin)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showMainNav(self):
        QMain = QMain(self)
        self.widget.addWidget(QMain)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showSettingsNav(self):
        self.QSettings = QSettings(self)
        self.QSettings.show()

    def showTechnicianForm(self):
        QTechnician = QTechnician(self)
        self.widget.addWidget(QTechnician)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)