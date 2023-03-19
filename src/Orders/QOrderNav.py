from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QIcon


class QOrderNav(QMainWindow):
    def __init__(self, model, view):
        super(QOrderNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Culture_Order_Forms_Nav.ui", self)
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.culture.clicked.connect(self.handleCulturePressed)
        self.duwl.clicked.connect(self.handleDUWLPressed)
        self.back.clicked.connect(self.handleBackPressed)

    ##@throwsViewableException
    def handleCulturePressed(self):
        self.view.showCultureOrderForm()
        self.close()

    ##@throwsViewableException
    def handleDUWLPressed(self):
        self.view.showDUWLNav()
        self.close()

    ##@throwsViewableException
    def handleBackPressed(self):
        self.close()