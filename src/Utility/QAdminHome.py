from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin
from Utility.QPrefixGraph import QPrefixGraph


class QAdminHome(QMainWindow):
    def __init__(self, model, view):
        super(QAdminHome, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Admin_Home_Screen.ui", self)
        self.settings.setIcon(QIcon('Icon/settingsIcon.png'))
        self.logout.setIcon(QIcon('Icon/logoutIcon.png'))
        self.cultureOrder.clicked.connect(self.handleCultureOrderFormsPressed)
        self.resultEntry.clicked.connect(self.handleResultEntryPressed)
        self.qaReport.clicked.connect(self.handleQAReportPressed)
        self.settings.clicked.connect(self.handleSettingsPressed)
        self.logout.clicked.connect(self.handleLogoutPressed)
        QPrefixGraph(self.model)
        #print(self.model.getCurrUser())
        #self.view.auditor(currentTech, "TEST", 'SAMPLE ID', 'TYPE')

    #@throwsViewableException
    def handleCultureOrderFormsPressed(self):
        self.view.showCultureOrderNav()

    #@throwsViewableException
    def handleResultEntryPressed(self):
        self.view.showResultEntryNav()

    ##@throwsViewableException
    def handleQAReportPressed(self):
        self.view.showQAReportScreen()

    #@throwsViewableException
    def handleSettingsPressed(self):
        self.view.showSettingsNav()

    #@throwsViewableException
    def handleLogoutPressed(self):
        self.view.showAdminLoginScreen()