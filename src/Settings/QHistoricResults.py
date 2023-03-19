from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon

class QHistoricResults(QMainWindow):
    def __init__(self, model, view):
        super(QHistoricResults, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Settings_Historical_Results_Form.ui", self)
        #self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.back.clicked.connect(self.handleBackPressed)
        #self.home.clicked.connect(self.handleReturnToMainMenuPressed)

    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()