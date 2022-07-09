from PyQt5.QtWidgets import QMainWindow
from PyQt5.uic import loadUi
import bcrypt

class QTechEditor(QMainWindow):
    def __init__(self, app, eid: int):
        super(QTechEditor, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Settings_Edit_Technician.ui', self)
        self.back.clicked.connect(self.close)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)

        fields = ('[Entry]', '[First]', '[Middle]', '[Last]', '[Username]', '[Password]')
        self.technician = app.db.select('Techs', fields, 1, eid)
        self.firstName.setText(self.technician[0])
        self.middleName.setText(self.technician[1])
        self.lastName.setText(self.technician[2])
        self.username.setText(self.technician[3])

        self.app = app
        self.eid = eid

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

    def handleSavePressed(self):
        if self.newPassword.text()==self.confirmNewPassword.text():
            if bcrypt.checkpw(self.oldPassword.text().encode('utf-8'), self.technician[4].encode('utf-8')):
                self.model.updateTech(
                    self.id,
                    self.firstName.text(),
                    self.middleName.text(),
                    self.lastName.text(),
                    self.username.text(),
                    self.newPassword.text()
                )
                self.close()
            else: self.view.showErrorScreen('Old password is incorrect')
        else: self.view.showErrorScreen('New password and confirm new password are mismatched')