from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QIcon

class QSettingsNav(QMainWindow):
    def __init__(self, model, view):
        super(QSettingsNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Admin_Settings_Nav.ui", self)
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.technicianSettings.clicked.connect(self.handleTechnicianSettingsPressed)
        self.manageArchives.clicked.connect(self.handleManageArchivesPressed)
        self.managePrefixes.clicked.connect(self.handleManagePrefixesPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.changeDatabase.clicked.connect(self.handleChangeDatabasePressed)
        #self.historicResults.clicked.connect(self.handleHistoricResultsPressed)
        self.rejectionLog.clicked.connect(self.handleRejectionLogPressed)

    #@throwsViewableException
    def handleChangeDatabasePressed(self):
        self.view.showSetFilePathScreen()
        self.close()

    #@throwsViewableException
    def handleTechnicianSettingsPressed(self):
        self.view.showSettingsManageTechnicianForm()
        self.close()

    #@throwsViewableException
    def handleManageArchivesPressed(self):
        self.view.showSettingsManageArchivesForm()
        self.close()

    #@throwsViewableException
    def handleManagePrefixesPressed(self):
        self.view.showSettingsManagePrefixesForm()
        self.close()

    ##@throwsViewableException
    #def handleHistoricResultsPressed(self):
        #self.view.showHistoricResultsForm()
        #self.close()

    #@throwsViewableException
    def handleRejectionLogPressed(self):
        self.view.showRejectionLogForm()
        self.close()
    
    #@throwsViewableException
    def handleBackPressed(self):
        self.close()