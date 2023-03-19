from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow

class QConfirmation(QMainWindow):
    def __init__(self, model, view):
        super(QConfirmation, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Confirmation_Window.ui", self)
        self.Cancel.clicked.connect(self.handleCancelPressed)

    #@throwsViewableException
    def handleCancelPressed(self):
        self.close()