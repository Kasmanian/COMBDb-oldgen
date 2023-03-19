from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow

class QError(QMainWindow):
    def __init__(self, model, view, message):
        super(QError, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Error_Window.ui", self)
        self.ok.clicked.connect(self.handleOKPressed)
        self.errorMessage.setText(str(message))

    #@throwsViewableException
    def handleOKPressed(self):
        self.close()