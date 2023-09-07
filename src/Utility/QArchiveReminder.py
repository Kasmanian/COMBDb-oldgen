from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow

from Utility.QViewableException import QViewableException

class QArchiveReminder(QMainWindow):
    def __init__(self, model, view):
        super(QArchiveReminder, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Archive_Prompt.ui", self)
        self.no.clicked.connect(self.handleNoPressed)
    
    @QViewableException.throwsViewableException
    def handleNoPressed(self):
        self.close()