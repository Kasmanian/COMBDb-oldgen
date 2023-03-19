from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow


class QArchiveReminder(QMainWindow):
    def __init__(self, model, view):
        super(QArchiveReminder, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Archive_Prompt.ui", self)
        self.no.clicked.connect(self.handleNoPressed)
    
    #@throwsViewableException
    def handleNoPressed(self):
        self.close()