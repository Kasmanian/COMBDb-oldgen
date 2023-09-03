from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QIcon

from Utility.QViewableException import QViewableException

class QManageArchives(QMainWindow): #TODO - incorporate archiving.
    def __init__(self, model, view):
        super(QManageArchives, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Settings_Manage_Archives_Form.ui", self)
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()

    @QViewableException.throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()