from PyQt5.uic import loadUi
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon

from Utility.QViewableException import QViewableException


class QAdminLogin(QMainWindow):
    def __init__(self, model, view):
        super(QAdminLogin, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Admin_Login.ui", self)
        self.login.setIcon(QIcon("Icon/loginIcon.png"))
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.login.clicked.connect(self.handleLoginPressed)

    @QViewableException.throwsViewableException
    def handleLoginPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if len(self.user.text()) == 0 or len(self.pswd.text()) == 0:
            self.errorMessage.setText("Please input all fields")
        else:
            if self.model.techLogin(self.user.text(), self.pswd.text()):
                global currentTech
                currentTech = list(self.model.currentTech(self.user.text(), "Entry"))[0]
                self.view.auditor(currentTech, "Login", "COMBDb", "System")
                self.view.showAdminHomeScreen()
            else:
                self.errorMessage.setText("Invalid username or password")

    @QViewableException.throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

    @QViewableException.throwsViewableException
    def event(self, event):
        if event.type() == QtCore.QEvent.KeyPress:
            if event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
                self.handleLoginPressed()
        return super().event(event)