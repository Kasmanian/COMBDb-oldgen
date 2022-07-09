from PyQt5.QtWidgets import QMainWindow
from PyQt5.uic import loadUi

class QConfirmation(QMainWindow):
    def __init__(self):
        super(QConfirmation, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Confirmation_Window.ui', self)
        self.Cancel.clicked.connect(self.close)