from PyQt5.QtWidgets import QMainWindow
from PyQt5.uic import loadUi

class QMain(QMainWindow):
    def __init__(self, app):
        super(QMain, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Admin_Home_Screen.ui', self)
        self.cultureOrder.clicked.connect(app.showCultureOrderNav)
        self.resultEntry.clicked.connect(app.showResultEntryNav)
        self.settings.clicked.connect(app.showSettingsNav)
        self.logout.clicked.connect(app.showAdminLoginScreen)

#close Qsettings!
class QSettings(QMainWindow):
    def __init__(self, app):
        super(QSettings, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Admin_Settings_Nav.ui', self)
        self.technicianSettings.clicked.connect(app.showSettingsManageTechnicianForm)
        self.manageArchives.clicked.connect(app.showSettingsManageArchivesForm)
        self.managePrefixes.clicked.connect(app.showSettingsManagePrefixesForm)
        self.changeDatabase.clicked.connect(app.showSetFilePathScreen)
        self.back.clicked.connect(self.close)