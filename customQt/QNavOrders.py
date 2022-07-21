from PyQt5.QtWidgets import QMainWindow
from PyQt5.uic import loadUi

class QOrders(QMainWindow):
    def __init__(self, app):
        super(QOrders, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Culture_Order_Forms_Nav.ui', self)
        self.culture.clicked.connect(self.handleCulturePressed)
        self.duwl.clicked.connect(self.handleDUWLPressed)
        self.back.clicked.connect(self.close)

        self.app = app

    def handleCulturePressed(self):
        self.app.showCultureOrderForm()
        self.close()

    def handleDUWLPressed(self):
        self.app.showDUWLNav()
        self.close()