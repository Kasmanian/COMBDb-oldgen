from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QIcon

from Utility.QViewableException import QViewableException

class QDUWLNav(QMainWindow):
    def __init__(self, model, view):
        super(QDUWLNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_DUWL_Nav.ui", self)
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.orderCulture.clicked.connect(self.handleOrderCulturePressed)
        self.receivingCulture.clicked.connect(self.handleReceivingCulturePressed)
        self.back.clicked.connect(self.handleBackPressed)

    @QViewableException.throwsViewableException
    def handleOrderCulturePressed(self):
        self.close()
        self.view.showDUWLOrderForm()

    @QViewableException.throwsViewableException
    def handleReceivingCulturePressed(self):
        self.close()
        self.view.showDUWLReceiveForm()

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.close()
        self.view.showCultureOrderNav()