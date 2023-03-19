from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin

class QManageTechnicians(QMainWindow):
    def __init__(self, model, view):
        super(QManageTechnicians, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Settings_Manage_Technicians_Form.ui", self)
        self.addTech.setIcon(QIcon('Icon/addClinicianIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.edit.clicked.connect(self.handleEditPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.techTable.itemSelectionChanged.connect(self.handleTechnicianSelected)
        self.activate.clicked.connect(self.handleActivatePressed)
        self.deactivate.clicked.connect(self.handleDeactivatePressed)
        self.addTech.clicked.connect(self.handleAddTechPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.confirmPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.activate.setEnabled(False)
        self.deactivate.setEnabled(False)
        self.edit.setEnabled(False)
        self.selectedTechnician = []
        self.updateTable()

    #@throwsViewableException
    def updateTable(self):
        techs = self.model.selectTechs('Entry, Username, Active')
        self.techTable.setRowCount(0)
        self.techTable.setRowCount(len(techs)) 
        self.techTable.setColumnCount(3)
        self.techTable.setColumnWidth(0,75)
        self.techTable.setColumnWidth(1,150)
        self.techTable.setColumnWidth(2,75)
        for i in range(0, len(techs)):
            self.techTable.setItem(i,0, QTableWidgetItem(str(techs[i][0])))
            self.techTable.setItem(i,1, QTableWidgetItem(techs[i][1]))
            self.techTable.setItem(i,2, QTableWidgetItem(techs[i][2]))

    #@throwsViewableException
    def handleEditPressed(self):
        if len(self.selectedTechnician)>0:
            self.view.showEditTechnician(self.selectedTechnician[1])

    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    #@throwsViewableException
    def handleTechnicianSelected(self):
        self.activate.setEnabled(True)
        self.deactivate.setEnabled(True)
        self.edit.setEnabled(True)
        self.selectedTechnician = [
            self.techTable.currentRow(), 
            int(self.techTable.item(self.techTable.currentRow(), 0).text()),
            self.techTable.item(self.techTable.currentRow(), 1).text(),
            self.techTable.item(self.techTable.currentRow(), 2).text(),
        ]
        self.tech.setText(self.techTable.item(self.techTable.currentRow(), 1).text())
    
    #@throwsViewableException
    def handleActivatePressed(self): #TODO - KEEP ADDING AUDIT LOG FUNCTIONALITY
        if self.selectedTechnician[3] != 'Yes':
            if self.model.toggleTech(self.selectedTechnician[1], 'Yes'):
                self.selectedTechnician[3] = 'Yes'
                self.techTable.item(self.selectedTechnician[0], 2).setText('Yes')
                self.view.auditor(QAdminLogin.currentTech, 'Activate', self.selectedTechnician[2], 'Settings_Edit_Technician')

    #@throwsViewableException
    def handleDeactivatePressed(self):
        print(QAdminLogin.currentTech)
        if self.selectedTechnician[3] != 'No':
            if self.model.toggleTech(self.selectedTechnician[1], 'No'):
                self.selectedTechnician[3] = 'No'
                self.techTable.item(self.selectedTechnician[0], 2).setText('No')
                self.view.auditor(QAdminLogin.currentTech, 'Deactivate', self.selectedTechnician[2], 'Settings_Edit_Technician')

    #@throwsViewableException
    def handleAddTechPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        self.techTable.clearSelection()
        user = self.user.text()
        if self.pswd.text()==self.confirmPswd.text() and self.pswd.text() and self.confirmPswd.text():
            if self.fName.text() and self.lName.text() and self.user.text():
                if self.model.findTechUsername(self.user.text()) == None:
                    self.model.addTech(self.fName.text(), self.mName.text(), self.lName.text(), self.user.text(), self.pswd.text())
                    self.updateTable()
                    self.handleClearPressed()
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Successfully added technician: " + user)
                    self.view.auditor(QAdminLogin.currentTech, 'Add', user, 'Settings_Edit_Technician')
                else:
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("A technician with this username already exists")
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("You must have a first name, last name, and username")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Password and confirm password are required and must match")

    #@throwsViewableException
    def handleClearPressed(self):
        self.fName.clear()
        self.mName.clear()
        self.lName.clear()
        self.user.clear()
        self.pswd.clear()
        self.confirmPswd.clear()
        self.techTable.clearSelection()
        self.tech.clear()
        self.activate.setEnabled(False)
        self.deactivate.setEnabled(False)
        self.edit.setEnabled(False)

    #@throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")