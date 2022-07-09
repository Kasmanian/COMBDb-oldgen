from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from PyQt5.uic import loadUi

class SettingsManageTechnicianForm(QMainWindow):
    def __init__(self, app):
        super(SettingsManageTechnicianForm, self).__init__()
        self.app = app
        loadUi('COMBDb/UI Screens/COMBdb_Settings_Manage_Technicians_Form.ui', self)
        self.edit.clicked.connect(self.handleEditPressed)
        self.back.clicked.connect(app.showSettingsNav)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.technicianTable.itemSelectionChanged.connect(self.handleTechnicianSelected)
        self.activate.clicked.connect(self.handleActivatePressed)
        self.deactivate.clicked.connect(self.handleDeactivatePressed)
        self.addTech.clicked.connect(self.handleAddTechPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.selectedTechnician = []
        self.updateTable()

    def updateTable(self):
        techs = self.model.selectTechs('Entry, Username, Active')
        self.technicianTable.setRowCount(0)
        self.technicianTable.setRowCount(len(techs)) 
        self.technicianTable.setColumnCount(3)
        try:
            for i in range(0, len(techs)):
                self.technicianTable.setItem(i,0, QTableWidgetItem(str(techs[i][0])))
                self.technicianTable.setItem(i,1, QTableWidgetItem(techs[i][1]))
                self.technicianTable.setItem(i,2, QTableWidgetItem(techs[i][2]))
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleEditPressed(self):
        if len(self.selectedTechnician)>0:
            self.app.showEditTechnician(self.selectedTechnician[1])

    def handleBackPressed(self):
        self.view.showSettingsNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleTechnicianSelected(self):
        self.selectedTechnician = [
            self.technicianTable.currentRow(), 
            int(self.technicianTable.item(self.technicianTable.currentRow(), 0).text()),
            self.technicianTable.item(self.technicianTable.currentRow(), 1).text(),
            self.technicianTable.item(self.technicianTable.currentRow(), 2).text(),
        ]
        self.technician.setText(self.technicianTable.item(self.technicianTable.currentRow(), 1).text())
    
    def handleActivatePressed(self):
        try:
            if self.selectedTechnician[3] != 'Yes':
                if self.model.toggleTech(self.selectedTechnician[1], 'Yes'):
                    self.selectedTechnician[3] = 'Yes'
                    self.technicianTable.item(self.selectedTechnician[0], 2).setText('Yes')
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleDeactivatePressed(self):
        try:
            if self.selectedTechnician[3] != 'No':
                if self.model.toggleTech(self.selectedTechnician[1], 'No'):
                    self.selectedTechnician[3] = 'No'
                    self.technicianTable.item(self.selectedTechnician[0], 2).setText('No')
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleAddTechPressed(self):
        try:
            if self.password.text()==self.confirmPassword.text():
                if self.firstName.text() and self.lastName.text() and self.username.text():
                    self.model.addTech(self.firstName.text(), self.middleName.text(), self.lastName.text(), self.username.text(), self.password.text())
                    self.updateTable()
                    self.handleClearPressed()
                else: self.view.showErrorMessage('You must have a first name, last name, and username')
            else: self.view.showErrorMessage('Password and confirm password must match')
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        self.firstName.clear()
        self.middleName.clear()
        self.lastName.clear()
        self.username.clear()
        self.password.clear()
        self.confirmPassword.clear()