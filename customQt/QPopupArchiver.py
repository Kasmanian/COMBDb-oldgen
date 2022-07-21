from PyQt5.QtWidgets import QMainWindow
from PyQt5.uic import loadUi

class QArchiver(QMainWindow):
    def __init__(self):
        super(QArchiver, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Archive_Prompt.ui', self)
        self.no.clicked.connect(self.close)