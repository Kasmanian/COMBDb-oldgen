from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon
from PyQt5 import QtWidgets
import bcrypt

from Utility.QViewableException import QViewableException


class QEditTechnician(QMainWindow):
    def __init__(self, model, view, id):
        super(QEditTechnician, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.id = id
        loadUi("UI Screens/COMBdb_Settings_Edit_Technician.ui", self)
        self.save.setIcon(QIcon("Icon/saveIcon.png"))
        self.home.setIcon(QIcon("Icon/menuIcon.png"))
        self.back.setIcon(QIcon("Icon/backIcon.png"))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.tech = self.model.findTech(
            self.id, "[First], [Middle], [Last], [Username], [Password]"
        )
        self.fName.setText(self.tech[0])
        self.mName.setText(self.tech[1])
        self.lName.setText(self.tech[2])
        self.user.setText(self.tech[3])
        self.oldPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.newPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.confirmNewPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.save.clicked.connect(self.handleSavePressed)

    @QViewableException.throwsViewableException
    def handleBackPressed(self):
        self.close()

    @QViewableException.throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

    @QViewableException.throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if (
            self.fName.text()
            and self.lName.text()
            and self.user.text()
            and self.oldPswd.text()
            and self.newPswd.text()
            and self.confirmNewPswd.text()
        ):
            if self.newPswd.text() == self.confirmNewPswd.text():
                if bcrypt.checkpw(
                    self.oldPswd.text().encode("utf-8"), self.tech[4].encode("utf-8")
                ):
                    self.model.updateTech(
                        self.id,
                        self.fName.text(),
                        self.mName.text(),
                        self.lName.text(),
                        self.user.text(),
                        self.newPswd.text(),
                    )
                    self.view.auditor(
                        self.view.currentTech,
                        "Edit",
                        self.user.text(),
                        "Settings_Edit_Technician",
                    )
                    self.close()
                else:
                    self.errorMessage.setStyleSheet(
                        "font: 12pt 'MS Shell Dlg 2'; color: red"
                    )
                    self.errorMessage.setText("Old password is incorrect")
            else:
                self.errorMessage.setStyleSheet(
                    "font: 12pt 'MS Shell Dlg 2'; color: red"
                )
                self.errorMessage.setText(
                    "New password and confirm new password don't match"
                )
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Missing required fields")

    @QViewableException.throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")