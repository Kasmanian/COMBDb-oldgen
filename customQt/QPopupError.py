from PyQt5.QtWidgets import QMainWindow
from PyQt5.uic import loadUi

class QError(QMainWindow):
    def __init__(self, message):
        super(QError, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Error_Window.ui', self)
        self.ok.clicked.connect(self.close)
        self.errorMessage.setText(str(message))