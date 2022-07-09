from PyQt5.QtWidgets import QMainWindow, QLineEdit
from PyQt5.uic import loadUi

class QLogin(QMainWindow):
    def __init__(self, app):
        super(QLogin, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Admin_Login.ui', self)
        self.login.clicked.connect(self.handleLoginPressed)
        self.pswd.setEchoMode(QLineEdit.Password)
        self.app = app

    def handleLoginPressed(self):
        if self.app.db.techLogin(self.usrnm.text(), self.pswd.text()):
            self.app.showAdminHomeScreen()