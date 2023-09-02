from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QIcon


class QResultNav(QMainWindow):
    def __init__(self, model, view):
        super(QResultNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Result_Entry_Forms_Nav.ui", self)
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.culture.clicked.connect(self.handleCulturePressed)
        self.cat.clicked.connect(self.handleCATPressed)
        self.duwl.clicked.connect(self.handleDUWLPressed)
        self.back.clicked.connect(self.handleBackPressed)

    #@throwsViewableException
    def handleCulturePressed(self):
        self.close()
        self.view.showCultureResultForm()

    #@throwsViewableException
    def handleCATPressed(self):
        self.close()
        self.view.showCATResultForm()

    #@throwsViewableException
    def handleDUWLPressed(self):
        self.close()
        self.view.showDUWLResultForm()

    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()
        self.close()