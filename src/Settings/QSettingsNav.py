from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QIcon

from Utility.QViewableException import QViewableException

class QSettingsNav(QMainWindow):
    def __init__(self, model, view):
        super(QSettingsNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Admin_Settings_Nav.ui", self)
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.technicianSettings.clicked.connect(self.handleTechnicianSettingsPressed)
        self.managePrefixes.clicked.connect(self.handleManagePrefixesPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.changeDatabase.clicked.connect(self.handleChangeDatabasePressed)
        self.rejectionLog.clicked.connect(self.handleRejectionLogPressed)

    @QViewableException.throwsViewableException
    def handleChangeDatabasePressed(self):
        self.view.showSetFilePathScreen()
        self.close()

    @QViewableException.throwsViewableException
    def handleTechnicianSettingsPressed(self):
        self.view.showSettingsManageTechnicianForm()
        self.close()

    @QViewableException.throwsViewableException
    def handleManagePrefixesPressed(self):
        self.view.showSettingsManagePrefixesForm()
        self.close()

    @QViewableException.throwsViewableException
    def handleRejectionLogPressed(self):
        self.view.showRejectionLogForm()
        self.close()
    
    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()
        self.close()