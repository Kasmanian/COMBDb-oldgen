from PyQt5.QtWidgets import QMainWindow
from PyQt5.uic import loadUi

class QMain(QMainWindow):
    def __init__(self, app):
        super(QMain, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Admin_Home_Screen.ui', self)
        self.cultureOrder.clicked.connect(app.showCultureOrderNav)
        self.resultEntry.clicked.connect(app.showResultEntryNav)
        self.settings.clicked.connect(app.showSettingsNav)
        self.logout.clicked.connect(app.showAdminLoginScreen)