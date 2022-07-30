from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets, QtPrintSupport
from PyQt5.QtWidgets import *
import win32com.client as win32
import sys, os, datetime, json
from mailmerge import MailMerge
from docxtpl import DocxTemplate
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtCore import QUrl, Qt, QDate, pyqtSignal
from PyQt5.QtGui import QIcon
import bcrypt

def passPrintPrompt(boolean):
        pass

class View:
    def __init__(self, model):
        self.model = model
        app = QApplication(sys.argv)
        app.setApplicationDisplayName('COMBDb')
        screen = AdminLoginScreen(model, self)
        self.widget = QtWidgets.QStackedWidget()
        self.widget.addWidget(screen)
        self.widget.setGeometry(10,10,1000,800)
        self.widget.showMaximized()
        if not self.model.connect():
            self.showSetFilePathScreen()
        else:
            self.setClinicianList()
        try:
            sys.exit(app.exec())
        except Exception as e:
            self.showErrorScreen(e)

    def showSetFilePathScreen(self):
        self.setFilePathScreen = SetFilePathScreen(self.model, self)
        self.setFilePathScreen.show()

    def showErrorScreen(self, message):
        self.setErrorScreen = SetErrorScreen(self.model, self, message)
        self.setErrorScreen.show()

    def showConfirmationScreen(self):
        self.setConfirmationScreen = SetConfirmationScreen(self.model, self)
        self.setConfirmationScreen.show()

    def showArchiveReminderScreen(self):
        self.setArchiveReminderScreen = SetArchiveReminderScreen(self.model, self)
        self.setArchiveReminderScreen.show()

    def showAdminLoginScreen(self):
        adminLoginScreen = AdminLoginScreen(self.model, self)
        self.widget.addWidget(adminLoginScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAdminHomeScreen(self):
        adminHomeScreen = AdminHomeScreen(self.model, self)
        self.widget.addWidget(adminHomeScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showSettingsNav(self):
        self.settingsNav = SettingsNav(self.model, self)
        self.settingsNav.show()

    def showSettingsManageTechnicianForm(self):
        settingsManageTechnicianForm = SettingsManageTechnicianForm(self.model, self)
        self.widget.addWidget(settingsManageTechnicianForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showEditTechnician(self, id):
        self.settingsEditTechnician = SettingsEditTechnician(self.model, self, id)
        self.settingsEditTechnician.show()

    def showSettingsManageArchivesForm(self):
        settingsManageArchivesForm = SettingsManageArchivesForm(self.model, self)
        self.widget.addWidget(settingsManageArchivesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showSettingsManagePrefixesForm(self):
        settingsManagePrefixesForm = SettingsManagePrefixesForm(self.model, self)
        self.widget.addWidget(settingsManagePrefixesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCultureOrderNav(self):
        self.cultureOrderNav = CultureOrderNav(self.model, self)
        self.cultureOrderNav.show()

    def showCultureOrderForm(self):
        cultureOrderForm = CultureOrderForm(self.model, self)
        self.widget.addWidget(cultureOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAddClinicianScreen(self, dropdown):
        self.addClinician = AddClinician(self.model, self, dropdown)
        self.addClinician.show()

    def showDUWLNav(self):
        self.duwlNav = DUWLNav(self.model, self)
        self.duwlNav.show()

    def showDUWLOrderForm(self):
        duwlOrderForm = DUWLOrderForm(self.model, self)
        self.widget.addWidget(duwlOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLReceiveForm(self):
        duwlReceiveForm = DUWLReceiveForm(self.model, self)
        self.widget.addWidget(duwlReceiveForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showResultEntryNav(self):
        self.resultEntryNav = ResultEntryNav(self.model, self)
        self.resultEntryNav.show()

    def showCultureResultForm(self):
        cultureResultForm = CultureResultForm(self.model, self)
        self.widget.addWidget(cultureResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCATResultForm(self):
        catResultForm = CATResultForm(self.model, self)
        self.widget.addWidget(catResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLResultForm(self):
        duwlResultForm = DUWLResultForm(self.model, self)
        self.widget.addWidget(duwlResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showPrintPreview(self, path):
        self.web = QWebEngineView()
        self.web.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)
        self.web.setWindowTitle('Print Preview')
        self.web.setContextMenuPolicy(Qt.ActionsContextMenu)
        printAction = QAction('Print', self.web)
        printAction.triggered.connect(self.showPrintPrompt)
        self.web.addAction(printAction)
        self.web.load(QUrl.fromLocalFile(path))
        self.web.showMaximized()

    def showPrintPrompt(self):
        self.dialog = QtPrintSupport.QPrintDialog()
        if self.dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.web.page().print(self.dialog.printer(), passPrintPrompt)

    def convertAndPrint(self, path):
        try:
            word = win32.DispatchEx('Word.Application')
            document = word.Documents.Open(path)
            tempPath = path.split('.')[0] + '.pdf'
            document.SaveAs(tempPath, 17)
            document.Close()
            # word.ActiveDocument()
            os.remove(path)
            word.Quit()
            self.showPrintPreview(tempPath)
        except Exception as e:
            self.showErrorScreen(e)

    def tempify(self, path):
        tempPath = path.split('\\')
        tempPath[len(tempPath)-1] = 'temp.docx'
        tempPath = '\\'.join(tempPath)
        return tempPath

    def fClinicianName(self, prefix, first, last, designation):
        em = ''
        comma = ', ' if first is not None else ''
        prefix = prefix+' ' if prefix is not None else prefix
        return f'{last or em}{comma}{prefix or em}{first or em}' if prefix is not None or first is not None or last is not None else designation or ''

    def fClinicianNameNormal(self, prefix, first, last, designation):
        em = ''
        prefix = prefix+' ' if prefix is not None else prefix
        first = first+' ' if first is not None else first
        return f'{prefix or em}{first or em}{last or em}' if prefix is not None or first is not None or last is not None else designation or ''

    def fSlashDate(self, date):
        if isinstance(date, datetime.datetime):
            return date.strftime('%m/%d/%Y')
        else:
            return f'{date.month()}/{date.day()}/{date.year()}'

    def dtToQDate(self, date):
        return QDate(date.year, date.month, date.day) if date is not None else QDate(self.model.date.year, self.model.date.month, self.model.date.day)

    def setClinicianList(self):
        try:
            self.clinicians = self.model.selectClinicians('Entry, Prefix, First, Last, Designation, Phone, Fax, Email, [Address 1], [Address 2], City, State, Zip, Enrolled, Inactive, Comments')
            self.entries = {}
            self.names = []
            for clinician in self.clinicians:
                name = self.fClinicianName(clinician[1], clinician[2], clinician[3], clinician[4])
                self.names.append(name)
                self.entries[name] = { 'db': clinician[0] }
            self.names.sort()
            for i in range(0, len(self.names)):
                self.entries[self.names[i]]['list'] = i
        except Exception as e:
            self.showErrorScreen(e)


class SetFilePathScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetFilePathScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Set_File_Path_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        self.browse.clicked.connect(self.handleBrowsePressed)
        self.save.clicked.connect(self.handleSavePressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()

    def handleBrowsePressed(self):
        fname = QFileDialog.getOpenFileName(self, 'Open File', 'C:', 'MS Access Files (*.accdb)')
        self.filePath.setText(fname[0])

    def handleSavePressed(self):
        try:
            with open('COMBDb\local.json', 'r+') as JSON:
                data = json.load(JSON)
                data['DBQ'] = str(Path(self.filePath.text()))
                JSON.seek(0)  # rewind
                json.dump(data, JSON)
                JSON.truncate()
            if not self.model.connect():
                self.view.showErrorScreen('Could not open database with the specified path.')
            else:
                self.view.setClinicianList()
                self.close()
        except Exception as e:
            self.view.showErrorScreen(e)

class SetErrorScreen(QMainWindow):
    def __init__(self, model, view, message):
        super(SetErrorScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Error_Window.ui", self)
        # Handle 'OK' button clicked
        self.ok.clicked.connect(self.handleOKPressed)
        print(message)
        self.errorMessage.setText(str(message))

    # Method for 'OK' button functionality
    def handleOKPressed(self):
        self.close()


class SetConfirmationScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetConfirmationScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Confirmation_Window.ui", self)
        # Handle 'Cancel' button clicked
        self.Cancel.clicked.connect(self.handleCancelPressed)

    # Method for 'OK' button functionality
    def handleCancelPressed(self):
        self.close()


class SetArchiveReminderScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetArchiveReminderScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Archive_Prompt.ui", self)
        # Handle 'No' button clicked
        self.no.clicked.connect(self.handleNoPressed)

    # Method for 'No' button functionality
    def handleNoPressed(self):
        self.close()


class AdminLoginScreen(QMainWindow):
    def __init__(self, model, view):
        super(AdminLoginScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Admin_Login.ui", self)
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.login.clicked.connect(self.handleLoginPressed)

    def handleLoginPressed(self):
        u = self.user.text()
        p = self.pswd.text()
        if len(u)==0 or len(p)==0:
            self.errorMessage.setText("Please input all fields")
        else:
            if self.model.techLogin(self.user.text(), self.pswd.text()):
                self.view.showAdminHomeScreen()
            else:
                self.errorMessage.setText("Invalid username or password")


class AdminHomeScreen(QMainWindow):
    def __init__(self, model, view):
        super(AdminHomeScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Admin_Home_Screen.ui", self)
        self.cultureOrder.clicked.connect(self.handleCultureOrderFormsPressed)
        self.resultEntry.clicked.connect(self.handleResultEntryPressed)
        self.settings.clicked.connect(self.handleSettingsPressed)
        self.logout.clicked.connect(self.handleLogoutPressed)

    def handleCultureOrderFormsPressed(self):
        self.view.showCultureOrderNav()

    def handleResultEntryPressed(self):
        self.view.showResultEntryNav()

    def handleSettingsPressed(self):
        self.view.showSettingsNav()

    def handleLogoutPressed(self):
        self.view.showAdminLoginScreen()

class SettingsNav(QMainWindow):
    def __init__(self, model, view):
        super(SettingsNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Admin_Settings_Nav.ui", self)
        self.technicianSettings.clicked.connect(self.handleTechnicianSettingsPressed)
        # Handle 'Manage Archives' button clicked
        self.manageArchives.clicked.connect(self.handleManageArchivesPressed)
        # Handle 'Manage Prefixes' button clicked
        self.managePrefixes.clicked.connect(self.handleManagePrefixesPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

        self.changeDatabase.clicked.connect(self.handleChangeDatabasePressed)

    def handleChangeDatabasePressed(self):
        self.view.showSetFilePathScreen()
        self.close()

    def handleTechnicianSettingsPressed(self):
        self.view.showSettingsManageTechnicianForm()
        self.close()

    # Method for 'Manage Archives' button functionality
    def handleManageArchivesPressed(self):
        self.view.showSettingsManageArchivesForm()
        self.close()

    # Method for 'Manage Prefixes' button functionality
    def handleManagePrefixesPressed(self):
        self.view.showSettingsManagePrefixesForm()
        self.close()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()

class SettingsManageTechnicianForm(QMainWindow):
    def __init__(self, model, view):
        super(SettingsManageTechnicianForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Settings_Manage_Technicians_Form.ui", self)
        # Handle 'Edit' button clicked
        self.edit.clicked.connect(self.handleEditPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.techTable.itemSelectionChanged.connect(self.handleTechnicianSelected)
        self.activate.clicked.connect(self.handleActivatePressed)
        self.deactivate.clicked.connect(self.handleDeactivatePressed)
        self.addTech.clicked.connect(self.handleAddTechPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.selectedTechnician = []
        self.updateTable()

    def updateTable(self):
        techs = self.model.selectTechs('Entry, Username, Active')
        self.techTable.setRowCount(0)
        self.techTable.setRowCount(len(techs)) 
        self.techTable.setColumnCount(3)
        try:
            for i in range(0, len(techs)):
                self.techTable.setItem(i,0, QTableWidgetItem(str(techs[i][0])))
                self.techTable.setItem(i,1, QTableWidgetItem(techs[i][1]))
                self.techTable.setItem(i,2, QTableWidgetItem(techs[i][2]))
        except Exception as e:
            self.view.showErrorScreen(e)

    # Method for 'Edit' button functionality
    def handleEditPressed(self):
        if len(self.selectedTechnician)>0:
            self.view.showEditTechnician(self.selectedTechnician[1])

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showSettingsNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleTechnicianSelected(self):
        self.selectedTechnician = [
            self.techTable.currentRow(), 
            int(self.techTable.item(self.techTable.currentRow(), 0).text()),
            self.techTable.item(self.techTable.currentRow(), 1).text(),
            self.techTable.item(self.techTable.currentRow(), 2).text(),
        ]
        self.tech.setText(self.techTable.item(self.techTable.currentRow(), 1).text())
    
    def handleActivatePressed(self):
        try:
            if self.selectedTechnician[3] != 'Yes':
                if self.model.toggleTech(self.selectedTechnician[1], 'Yes'):
                    self.selectedTechnician[3] = 'Yes'
                    self.techTable.item(self.selectedTechnician[0], 2).setText('Yes')
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleDeactivatePressed(self):
        try:
            if self.selectedTechnician[3] != 'No':
                if self.model.toggleTech(self.selectedTechnician[1], 'No'):
                    self.selectedTechnician[3] = 'No'
                    self.techTable.item(self.selectedTechnician[0], 2).setText('No')
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleAddTechPressed(self):
        try:
            user = self.user.text()
            if self.pswd.text()==self.confirmPswd.text() and self.pswd.text() and self.confirmPswd.text():
                if self.fName.text() and self.lName.text() and self.user.text():
                    self.model.addTech(self.fName.text(), self.mName.text(), self.lName.text(), self.user.text(), self.pswd.text())
                    self.updateTable()
                    self.handleClearPressed()
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Successfully added technician: " + user)
                else:
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("You must have a first name, last name, and username")
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Password and confirm password are required and must match")
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        self.fName.clear()
        self.mName.clear()
        self.lName.clear()
        self.user.clear()
        self.pswd.clear()
        self.confirmPswd.clear()


class SettingsEditTechnician(QMainWindow):
    # Class for the Edit Technician UI
    def __init__(self, model, view, id):
        super(SettingsEditTechnician, self).__init__()
        self.view = view
        self.model = model
        self.id = id
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Settings_Edit_Technician.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.tech = self.model.findTech(self.id, '[First], [Middle], [Last], [Username], [Password]')
        self.fName.setText(self.tech[0])
        self.mName.setText(self.tech[1])
        self.lName.setText(self.tech[2])
        self.user.setText(self.tech[3])
        self.save.clicked.connect(self.handleSavePressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

    def handleSavePressed(self):
        if self.fName.text() and self.lName.text() and self.user.text() and self.oldPswd.text() and self.newPswd.text() and self.confirmNewPswd.text():
            if self.newPswd.text()==self.confirmNewPswd.text():
                if bcrypt.checkpw(self.oldPswd.text().encode('utf-8'), self.tech[4].encode('utf-8')):
                    self.model.updateTech(
                        self.id,
                        self.fName.text(),
                        self.mName.text(),
                        self.lName.text(),
                        self.user.text(),
                        self.newPswd.text()
                    )
                    self.close()
                else: 
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText('Old password is incorrect')
            else: 
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText('New password and confirm new password are mismatched')
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText('Missing required fields')

class SettingsManageArchivesForm(QMainWindow): #TODO - incorporate archiving.
    # Class for the Manage Archives UI
    def __init__(self, model, view):
        super(SettingsManageArchivesForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Settings_Manage_Archives_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)

    def handleBackPressed(self):
        self.view.showSettingsNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    
class SettingsManagePrefixesForm(QMainWindow): #TODO - Need to populate tables and allow them to be edited.
    # Class for the Manage Prefixes UI
    def __init__(self, model, view):
        super(SettingsManagePrefixesForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Settings_Manage_Prefixes_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Return to Main Menu' button clicked
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.add.clicked.connect(self.handleAddPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)

        self.aeTWid.itemSelectionChanged.connect(lambda: self.handlePrefixSelected("Aerobic"))
        self.anTWid.itemSelectionChanged.connect(lambda: self.handlePrefixSelected("Anaerobic"))
        self.abTWid.itemSelectionChanged.connect(lambda: self.handlePrefixSelected("Antibiotic"))

        self.currentPrefix = ""
        self.selectedPrefix = {}
        self.updateTable("Aerobic")
        self.updateTable("Anaerobic")
        self.updateTable("Antibiotics")

    def updateTable(self, type):
        widget = self.aeTWid if type == "Aerobic" else self.anTWid if type == "Anaerobic" else self.abTWid
        prefix = self.model.selectPrefixes(type, 'Prefix, Word')
        widget.setRowCount(0)
        widget.setRowCount(len(prefix))
        widget.setColumnCount(2)
        widget.setColumnWidth(0, 50)
        widget.setColumnWidth(1, 300)
        try:
            for i in range(0, len(prefix)):
                widget.setItem(i,0, QTableWidgetItem(prefix[i][0]))
                widget.setItem(i,1, QTableWidgetItem(prefix[i][1]))
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePrefixSelected(self, type):
        widget = self.aeTWid if type == "Aerobic" else self.anTWid if type == "Anaerobic" else self.abTWid
        #print(widget.item(widget.currentRow(), 0))
        #print(widget.item(widget.currentRow(), 1))
        prefix = widget.item(widget.currentRow(), 0)
        word = widget.item(widget.currentRow(), 1)
        if prefix and word:
            self.selectedPrefix = {prefix.text() : [type, word.text()]}
            self.pName.setText(list(self.selectedPrefix.keys())[0])
            keyList = self.selectedPrefix.get(list(self.selectedPrefix.keys())[0])
            self.type.setCurrentIndex(self.type.findText(keyList[0]))
            self.word.setText(keyList[1])
            self.currentPrefix = self.model.findPrefix(self.pName.text(), 'Entry, Type, Prefix, Word')
            #print(self.currentPrefix)

    def handleBackPressed(self):
        self.view.showSettingsNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleAddPressed(self):   
        return

    def handleSavePressed(self):
        if self.pName.text() and self.word.text() and self.type.currentText():
            #print(self.currentPrefix[0])
            #print(self.type.currentText())
            #print(self.pName.text())
            #print(self.word.text())
            self.model.updatePrefixes(
                self.currentPrefix[0],
                self.type.currentText(),
                self.pName.text(),
                self.word.text()
            )
            self.updateTable(self.type.currentText())
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
            self.errorMessage.setText("Successfully Updated Prefix")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Type, Prefix and Word are required")

    def handleClearPressed(self):
        self.type.setCurrentIndex(0)
        self.pName.clear()
        self.word.clear()
        self.errorMessage.clear()


class CultureOrderNav(QMainWindow):
    def __init__(self, model, view):
        super(CultureOrderNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Forms_Nav.ui", self)
        self.culture.clicked.connect(self.handleCulturePressed)
        self.duwl.clicked.connect(self.handleDUWLPressed)
        self.back.clicked.connect(self.handleBackPressed)

    def handleCulturePressed(self):
        self.view.showCultureOrderForm()
        self.close()

    def handleDUWLPressed(self):
        self.view.showDUWLNav()
        self.close()

    def handleBackPressed(self):
        self.close()


class CultureOrderForm(QMainWindow):
    def __init__(self, model, view):
        super(CultureOrderForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Form.ui", self)
        self.find.setIcon(QIcon('COMBDb/Icon/searchIcon.png'))
        self.clinDrop.clear()
        #self.clinDrop.addItem(" ")
        self.clinDrop.addItems(self.view.names)
        self.addClinician.clicked.connect(self.handleAddNewClinicianPressed)
        self.find.clicked.connect(self.handleSearchPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.print.setEnabled(False)
        self.colDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))

    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    def handleSearchPressed(self):
        try:
            if not self.saID.text().isdigit():
                self.saID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('Cultures', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes]')
            if self.sample is None:
                self.sample = self.model.findSample('CATs', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes]')
                if self.sample is None:
                    self.saID.setText('xxxxxx')
            if self.sample is not None:
                self.chID.setText(self.sample[0])
                clinician = self.model.findClinician(self.sample[1])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.fName.setText(self.sample[2])
                self.lName.setText(self.sample[3])
                self.type.setCurrentIndex(self.type.findText(self.sample[4]))
                self.colDate.setDate(self.view.dtToQDate(self.sample[5]))
                self.recDate.setDate(self.view.dtToQDate(self.sample[6]))
                self.cText.setText(self.sample[7])
                self.nText.setText(self.sample[8])
                self.print.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
    
    def handleSavePressed(self):
        try:
            if self.fName.text() and self.lName.text() and self.type.currentText() and self.clinDrop.currentText() != " ":
                table = 'CATs' if self.type.currentText()=='Caries' else 'Cultures'
                saID = self.view.model.addPatientOrder(
                    table,
                    self.chID.text(),
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.fName.text(),
                    self.lName.text(),
                    self.colDate.date(),
                    self.recDate.date(),
                    self.type.currentText(),
                    self.cText.toPlainText(),
                    self.nText.toPlainText()
                )
                if saID:
                    self.saID.setText(str(saID))
                    self.save.setEnabled(False)
                    self.print.setEnabled(True)
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Successfully saved order: " + str(saID))
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter in all required fields")
        except Exception as e:
            self.view.showErrorScreen(e)
    
    def handlePrintPressed(self):
        try:
            if self.type.currentText()!='Caries':
                #print(f'clinician: {self.clinDrop.currentText()}')
                template = str(Path().resolve())+r'\COMBDb\templates\culture_worksheet_template.docx'
                dst = self.view.tempify(template)
                document = MailMerge(template)
                clinician=self.clinDrop.currentText().split(', ')
                document.merge(
                    saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:]}',
                    received=self.recDate.date().toString(),
                    type=self.type.currentText(),
                    chartID=self.chID.text(),
                    clinicianName = clinician[1] + " " + clinician[0],
                    patientName=f'{self.lName.text()}, {self.fName.text()}',
                    comments=self.cText.toPlainText(),
                    notes=self.nText.toPlainText()
                )
                document.write(dst)
                try:
                    self.view.convertAndPrint(dst)
                except Exception as e:
                    self.view.showErrorScreen(e)
            else:
                template = str(Path().resolve())+r'\COMBDb\templates\cat_worksheet_template.docx'
                dst = self.view.tempify(template)
                document = MailMerge(template)
                clinician=self.clinDrop.currentText().split(', ')
                document.merge(
                    saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:]}',
                    received=self.recDate.date().toString(),
                    chartID=self.chID.text(),
                    clinicianName = clinician[1] + " " + clinician[0],
                    patientName=f'{self.lName.text()}, {self.fName.text()}',
                )
                document.write(dst)
                try:
                    self.view.convertAndPrint(dst)
                except Exception as e:
                    self.view.showErrorScreen(e)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.fName.clear()
            self.lName.clear()
            self.colDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.saID.clear()
            self.chID.clear()
            self.cText.clear()
            self.nText.clear()
            self.clinDrop.setCurrentIndex(0)
            self.type.setCurrentIndex(0)
            self.save.setEnabled(True)
            self.print.setEnabled(False)
            self.clear.setEnabled(True)
            self.errorMessage.setText("")
        except Exception as e:
            self.view.showErrorScreen(e)

class AddClinician(QMainWindow):
    def __init__(self, model, view, dropdown):
        super(AddClinician, self).__init__()
        self.view = view
        self.model = model
        self.dropdown = dropdown
        loadUi("COMBDb/UI Screens/COMBdb_Add_New_Clinician.ui", self)
        self.clear.clicked.connect(self.handleClearPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.enrollDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))

    def handleSavePressed(self): #Incorporate validation to make sure clinician is actually added to DB
        try:
            if self.fName.text() and self.lName.text() and self.address1.text() and self.city.text() and self.state.currentText() and self.zip.text():
                self.model.addClinician(
                    self.title.currentText(),
                    self.fName.text(),
                    self.lName.text(),
                    self.designation.text(),
                    self.phone.text(),
                    self.fax.text(),
                    self.email.text(),
                    self.address1.text(),
                    self.address2.text(),
                    self.city.text(),
                    self.state.currentText(),
                    self.zip.text(),
                    None,
                    None,
                    self.cText.toPlainText()
                )
                self.view.setClinicianList()
                self.dropdown.clear()
                self.dropdown.addItems(self.view.names)
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("New clinician added: " + self.title.currentText() + " " +self.fName.text() + " " + self.lName.text())
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("First Name, Last Name, Address 1, City, State and Zip are all required")
        except Exception as e:
            self.view.showErrorScreen(e)
        #finally:
            #self.close()

    def handleBackPressed(self):
        self.close()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

    def handleClearPressed(self):
        try:
            self.title.setCurrentIndex(0)
            self.fName.clear()
            self.lName.clear()
            self.address1.clear()
            self.address2.clear()
            self.city.clear()
            self.state.setCurrentIndex(0)
            self.zip.clear()
            self.phone.clear()
            self.fax.clear()
            self.email.clear()
            self.enrollDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.designation.clear()
            self.cText.clear()
            self.errorMessage.clear()
        except Exception as e:
            self.view.showErrorScreen(e)

class DUWLNav(QMainWindow):
    def __init__(self, model, view):
        super(DUWLNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Nav.ui", self)
        self.orderCulture.clicked.connect(self.handleOrderCulturePressed)
        self.receivingCulture.clicked.connect(self.handleReceivingCulturePressed)
        self.back.clicked.connect(self.handleBackPressed)

    def handleOrderCulturePressed(self):
        self.close()
        self.view.showDUWLOrderForm()

    def handleReceivingCulturePressed(self):
        self.close()
        self.view.showDUWLReceiveForm()

    def handleBackPressed(self):
        self.close()
        self.view.showCultureOrderNav()

class DUWLOrderForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLOrderForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Order_Form.ui", self)
        self.find.setIcon(QIcon('COMBDb/Icon/searchIcon.png'))
        self.currentKit = 1
        self.kitList = []
        self.printList = {}
        self.kitNum.setText('1')
        self.numOrders.setValue(1)
        self.clinDrop.clear()
        self.clinDrop.addItems(self.view.names)
        self.shipDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.find.clicked.connect(self.handleSearchPressed)
        self.addClinician.clicked.connect(self.handleAddClinicianPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.clearAll.clicked.connect(self.handleClearAllPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.remove.clicked.connect(self.handleRemovePressed)
        self.kitTWid.setColumnCount(1)
        self.kitTWid.itemClicked.connect(self.activateRemove)
        self.print.setEnabled(False)
        self.remove.setEnabled(False)

    def activateRemove(self):
        self.remove.setEnabled(True)

    def handleSearchPressed(self):
        pass
        """
        try:
            pass
            if not self.saID.text().isdigit():
                self.saID.setText("xxxxxx")
            self.sample = self.model.findSample('Waterlines', int(self.saID.text()), 'Clinician, Comments, Notes, Shipped')
            if self.sample is None:
                self.saID.setText('xxxxxx')
            else:
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.cText.setText(self.sample[1])
                self.nText.setText(self.sample[2])
                self.shipDate.setDate(self.view.dtToQDate(self.sample[3]))
 

                self.kitList.append({
                    'sampleID': f'{str(self.saID)[0:2]}-{str(self.saID)[2:]}',
                    'clinician': 'Clinician___________________________',
                    'operatory': 'Operatory__________________________',
                    'collected': 'Collection Date______________________',
                    'clngagent': 'Cleaning Agent______________________'
                })
                self.printList[str(self.saID)] = self.currentKit-1
                self.currentKit += 1
                self.updateTable()
                self.save.setEnabled(False)
        except Exception as e:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText(e)
            self.view.showErrorScreen(e)
            """

    def handleAddClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleSavePressed(self):
        try:
            if self.clinDrop.currentText():
                numOrders = 1 if int(self.numOrders.text()) == None else int(self.numOrders.text())
                for x in range(numOrders):
                    saID = self.view.model.addWaterlineOrder(
                        self.view.entries[self.clinDrop.currentText()]['db'],
                        self.shipDate.date(),
                        self.cText.toPlainText(),
                        self.nText.toPlainText()
                    )
                    if saID: 
                        self.saID.setText(str(saID))
                        self.kitList.append({
                            'sampleID': f'{str(saID)[0:2]}-{str(saID)[2:]}',
                            'clinician': 'Clinician___________________________',
                            'operatory': 'Operatory__________________________',
                            'collected': 'Collection Date______________________',
                            'clngagent': 'Cleaning Agent______________________'
                        })
                        self.printList[str(saID)] = self.currentKit-1
                        self.currentKit += 1
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Created New DUWL Order: " + str(saID))
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please select a clinician")
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.kitNum.setText(str(self.currentKit))
            self.saID.setText('xxxxxx')
            self.cText.clear()
            self.nText.clear()
            self.numOrders.setValue(1)
            self.save.setEnabled(True)
            self.clear.setEnabled(True)
            self.clinDrop.setCurrentIndex(0)
            self.errorMessage.setText(" ")
            self.updateTable()
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearAllPressed(self):
        try:
            self.kitList.clear()
            self.currentKit = 1
            self.printList.clear()
            self.updateTable()
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleRemovePressed(self):
        try:
            del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
            del self.printList[self.kitTWid.currentItem().text()]
            count = 0
            for key in self.printList.keys():
                self.printList[key] = count
                count += 1
            self.updateTable()
            self.remove.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def updateTable(self):
        try:
            self.kitTWid.setRowCount(len(self.printList.keys()))
            count = 0
            for item in self.printList.keys():
                self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
                count += 1
            if len(self.printList.keys())>0:
                self.print.setEnabled(True)
            else:
                self.print.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePrintPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\duwl_label_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge_rows('sampleID', self.kitList)
            document.write(dst)
        except Exception as e:
            self.view.showErrorScreen(e)
        try:
            self.view.convertAndPrint(dst)
        except Exception as e:
            self.view.showErrorScreen(e)

class DUWLReceiveForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLReceiveForm, self).__init__()
        try:
            self.view = view
            self.model = model
            loadUi("COMBDb/UI Screens/COMBdb_DUWL_Receive_Form.ui", self)
            self.find.setIcon(QIcon('COMBDb/Icon/searchIcon.png'))
            self.clinDrop.clear()
            self.clinDrop.addItems(self.view.names)
            self.currentKit = 1
            self.kitList = []
            self.printList = {}
            self.save.setEnabled(False)
            self.print.setEnabled(False)
            self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.colDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.back.clicked.connect(self.handleBackPressed)
            self.home.clicked.connect(self.handleReturnToMainMenuPressed)
            self.save.clicked.connect(self.handleSavePressed)
            self.clear.clicked.connect(self.handleClearPressed)
            self.print.clicked.connect(self.handlePrintPressed)
            self.find.clicked.connect(self.handleSearchPressed)
            self.clearAll.clicked.connect(self.handleClearAllPressed)
            self.remove.clicked.connect(self.handleRemovePressed)
            self.kitTWid.setColumnCount(1)
            self.kitTWid.itemClicked.connect(self.activateRemove)
            self.print.setEnabled(False)
            self.remove.setEnabled(False)
        except Exception as e:
            print(e)

    def activateRemove(self):
        self.remove.setEnabled(True)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleSearchPressed(self):
        try:
            if not self.saID.text().isdigit():
                self.saID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('Waterlines', int(self.saID.text()), 'Clinician, Comments, Notes, OperatoryID, Product, Procedure, Collected, Received')
            if self.sample is None:
                self.saID.setText('xxxxxx')
            else:
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.cText.setText(self.sample[1])
                self.nText.setText(self.sample[2])
                self.operatory.setText(self.sample[3])
                self.product.setText(self.sample[4])
                self.procedure.setText(self.sample[5])
                self.colDate.setDate(self.view.dtToQDate(self.sample[6]))
                self.recDate.setDate(self.view.dtToQDate(self.sample[7]))
                self.save.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleSavePressed(self):
        try:
            if self.clinDrop.currentText():
                saID = int(self.saID.text())
                if self.model.addWaterlineReceiving(
                    saID,
                    self.operatory.text(),
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.colDate.date(),
                    self.recDate.date(),
                    self.product.text(),
                    self.procedure.text(),
                    self.cText.toPlainText(),
                    self.nText.toPlainText()
                ):
                    clinician = self.clinDrop.currentText().split(', ')
                    self.kitList.append({
                        'underline1': '__________',
                        'clinicianName': clinician[1] + " " + clinician[0],
                        'sampleID': f'{str(saID)[0:2]}-{str(saID)[2:]}',
                        'underline2': '__________',
                        'underline3': '__________'
                    })
                    self.printList[str(saID)] = self.currentKit-1
                    self.currentKit += 1
                    self.handleClearPressed()
                    self.save.setEnabled(False)
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Added DUWL Order: " + str(saID))
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please select a clinician")
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.saID.clear()
            self.clinDrop.setCurrentIndex(0)
            self.cText.clear()
            self.nText.clear()
            self.operatory.clear()
            self.procedure.clear()
            self.product.clear()
            self.save.setEnabled(True)
            self.clear.setEnabled(True)
            self.updateTable()
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearAllPressed(self):
        try:
            self.kitList.clear()
            self.currentKit = 1
            self.printList.clear()
            self.updateTable()
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleRemovePressed(self):
        try:
            del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
            del self.printList[self.kitTWid.currentItem().text()]
            count = 0
            for key in self.printList.keys():
                self.printList[key] = count
                count += 1
            self.updateTable()
            self.remove.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def updateTable(self):
        try:
            self.kitTWid.setRowCount(len(self.printList.keys()))
            count = 0
            for item in self.printList.keys():
                self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
                count += 1
            if len(self.printList.keys())>0:
                self.print.setEnabled(True)
            else:
                self.print.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePrintPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\pending_duwl_cultures_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge_rows('sampleID', self.kitList)
            document.merge(received=self.view.fSlashDate(self.recDate.date()))
            document.write(dst)
        except Exception as e:
            self.view.showErrorScreen(e)
        try:
            self.view.convertAndPrint(dst)
        except Exception as e:
            self.view.showErrorScreen(e)

    # def handleBackPressed(self):
    #     self.view.showCultureOrderNav()

    # def handleReturnToMainMenuPressed(self):
    #     self.view.showAdminHomeScreen()

class ResultEntryNav(QMainWindow):
    def __init__(self, model, view):
        super(ResultEntryNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Result_Entry_Forms_Nav.ui", self)
        self.culture.clicked.connect(self.handleCulturePressed)
        self.cat.clicked.connect(self.handleCATPressed)
        self.duwl.clicked.connect(self.handleDUWLPressed)
        self.back.clicked.connect(self.handleBackPressed)

    def handleCulturePressed(self):
        self.close()
        self.view.showCultureResultForm()

    def handleCATPressed(self):
        self.close()
        self.view.showCATResultForm()

    def handleDUWLPressed(self):
        self.close()
        self.view.showDUWLResultForm()

    def handleBackPressed(self):
        self.close()

class CultureResultForm(QMainWindow):
    def __init__(self, model, view):
        super(CultureResultForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Result_Form.ui", self)
        self.find.setIcon(QIcon('COMBDb/Icon/searchIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItems(self.view.names)
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.repDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.clear.clicked.connect(self.handleClearPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.find.clicked.connect(self.handleSearchPressed)
        self.printS.clicked.connect(self.handleDirectSmearPressed)
        self.printP.clicked.connect(self.handlePreliminaryPressed)
        self.printF.clicked.connect(self.handlePerioPressed)
        self.save.setEnabled(False)
        self.printP.setEnabled(False)
        self.printF.setEnabled(False)
        self.printS.setEnabled(False)
        #self.tableWidget.itemSelectionChanged.connect(self.handleCellChanged)
        #testbox = QComboBox()
        #self.tableWidget.setCellWidget(0, 0, testbox)
        self.anTWid.setRowCount(0)
        self.anTWid.setColumnCount(0)
        try:
            with open('COMBDb\local.json', 'r+') as JSON:
                count = 0
                data = json.load(JSON)
                self.aerobicPrefixes = data['PrefixToAerobic']
                self.aerobicBacteria = {}
                self.aerobicList = self.aerobicPrefixes.values()
                self.aerobicIndex = {}
                for prefix in self.aerobicPrefixes.keys():
                    self.aerobicBacteria[self.aerobicPrefixes[prefix]] = prefix
                    self.aerobicIndex[self.aerobicPrefixes[prefix]] = count
                    count += 1
                count = 0
                self.anaerobicPrefixes = data['PrefixToAnaerobic']
                self.anaerobicBacteria = {}
                self.anaerobicList = self.anaerobicPrefixes.values()
                self.anaerobicIndex = {}
                for prefix in self.anaerobicPrefixes.keys():
                    self.anaerobicBacteria[self.anaerobicPrefixes[prefix]] = prefix
                    self.anaerobicIndex[self.anaerobicPrefixes[prefix]] = count
                    count += 1
                count = 0
                self.antibioticPrefixes = data['PrefixToAntibiotics']
                self.antibiotics = {}
                self.antibioticsList = self.antibioticPrefixes.values()
                self.antibioticsIndex = {}
                for prefix in self.antibioticPrefixes.keys():
                    self.antibiotics[self.antibioticPrefixes[prefix]] = prefix
                    self.anaerobicIndex[prefix] = count
                    count += 1
                self.blacList = data['PrefixToB-Lac'].keys()
                self.growthList = data['PrefixToGrowth'].keys()
                self.susceptibilityList = data['PrefixToSusceptibility'].keys()
                self.headers = ['Growth', 'B-lac']
                self.headerIndexes = { 'Growth': 0, 'B-lac': 1 }
                self.options = ['NA'] + list(self.growthList) + list(self.blacList) + list(self.susceptibilityList)
                self.optionIndexes = { 'NA': 0, 'NI': 1, 'L': 2, 'M': 3, 'H': 4, 'P': 5, 'N': 6, 'S': 7, 'I': 8, 'R': 9 }
                for antibiotics in self.antibioticPrefixes.keys():
                    self.headers.append(antibiotics)
                    self.headerIndexes[antibiotics] = len(self.headers)-1
            self.addRow1.clicked.connect(self.addRowAerobic)
            self.addRow2.clicked.connect(self.addRowAnaerobic)
            self.delRow1.clicked.connect(self.delRowAerobic)
            self.delRow2.clicked.connect(self.delRowAnaerobic)
            self.addCol1.clicked.connect(self.addColAerobic)
            self.addCol2.clicked.connect(self.addColAnaerobic)
            self.delCol1.clicked.connect(self.delColAerobic)
            self.delCol2.clicked.connect(self.delColAnaerobic)
            self.aerobicTable = self.resultToTable(None)
            self.anaerobicTable = self.resultToTable(None)
            self.initTables()
            self.save.clicked.connect(self.handleSavePressed)
        except Exception as e:
            self.view.showErrorScreen(e)

    def initTables(self):
        try:
            self.aeTWid.setRowCount(0)
            self.aeTWid.setRowCount(len(self.aerobicTable))
            self.anTWid.setRowCount(0)
            self.anTWid.setRowCount(len(self.anaerobicTable))
            self.aeTWid.setColumnCount(0)
            self.aeTWid.setColumnCount(len(self.aerobicTable[0]))
            self.anTWid.setColumnCount(0)
            self.anTWid.setColumnCount(len(self.anaerobicTable[0]))
            self.aeTWid.setColumnWidth(0,300)
            self.anTWid.setColumnWidth(0,300)
            #aerobic
            self.aeTWid.setItem(0,0, QTableWidgetItem('Bacteria'))
            for i in range(0, len(self.aerobicTable)):
                for j in range(0, len(self.aerobicTable[0])):
                    item = IndexedComboBox(i, j, self, True)
                    if i>0 and j>0:
                        item.addItems(self.options)
                        item.setCurrentIndex(self.optionIndexes[self.aerobicTable[i][j]])
                    elif i<1 and j>0:
                        item.addItems(self.headers)
                        item.setCurrentIndex(self.headerIndexes[self.aerobicTable[i][j]])
                    elif i>0 and j<1:
                        item.addItems(self.aerobicList)
                        item.setCurrentIndex(self.aerobicIndex[self.aerobicTable[i][j]])
                    else: continue
                    self.aeTWid.setCellWidget(i, j, item)

            #anaerobic
            self.anTWid.setItem(0,0, QTableWidgetItem('Bacteria'))
            for i in range(0, len(self.anaerobicTable)):
                for j in range(0, len(self.anaerobicTable[0])):
                    item = IndexedComboBox(i, j, self, False)
                    if i>0 and j>0:
                        item.addItems(self.options)
                        item.setCurrentIndex(self.optionIndexes[self.anaerobicTable[i][j]])
                    elif i<1 and j>0:
                        item.addItems(self.headers)
                        item.setCurrentIndex(self.headerIndexes[self.anaerobicTable[i][j]])
                    elif i>0 and j<1:
                        item.addItems(self.anaerobicList)
                        item.setCurrentIndex(self.anaerobicIndex[self.anaerobicTable[i][j]])
                    else: continue
                    self.anTWid.setCellWidget(i, j, item)
        except Exception as e:
            self.view.showErrorScreen(e)

    #def eventFilter(self, source, event): trying to resolve the problem of mousehover/scrollwheel changing value in cells - shouldn't be allowed - solution is eventFilter, just need to figure it out
        #if source == self.
        #return

    def updateTable(self, kind, row, column):
        try:
            if kind:
                self.aerobicTable[row][column] = self.aeTWid.cellWidget(row, column).currentText() if self.aeTWid.cellWidget(row, column) else self.aerobicTable[row][column]
            else:
                self.anaerobicTable[row][column] = self.anTWid.cellWidget(row, column).currentText() if self.anTWid.cellWidget(row, column) else self.anaerobicTable[row][column]
        except Exception as e:
            self.view.showErrorScreen(e)

    def resultToTable(self, result):
        if result is not None:
            result = result.split('/')
            table = [[]]
            for i in range(0, len(result)):
                headers = ['Bacteria']
                bacteria = result[i].split(':')
                table.append([bacteria[0]])
                antibiotics = bacteria[1].split(';')
                for j in range(0, len(antibiotics)):
                    measures = antibiotics[j].split('=')
                    if i<1: headers.append(measures[0])
                    table[i+1].append(measures[1])
                if i<1: table[0] = headers
            return table
        else:
            return [['Bacteria','Growth', 'B-lac', 'PEN', 'AMP', 'CC', 'TET', 'CEP', 'ERY']]

    def tableToResult(self, table):
        if len(table)>1 and len(table[0])>1:
            result = ''
            for i in range(1, len(table)):
                if i>1: result += '/'
                result += f'{table[i][0]}:'
                for j in range(1, len(table[i])):
                    if j>1: result += ';'
                    result += f'{table[0][j]}={table[i][j]}'
            return result
        else:
            return None

    def addRowAerobic(self):
        try:
            self.aeTWid.setRowCount(self.aeTWid.rowCount()+1)
            self.aerobicTable.append(['Alpha-Hemolytic Streptococcus'])
            bacteria = IndexedComboBox(self.aeTWid.rowCount()-1, 0, self, True)
            bacteria.addItems(self.aerobicList)
            self.aeTWid.setCellWidget(self.aeTWid.rowCount()-1, 0, bacteria)
            for i in range(1, self.aeTWid.columnCount()):
                self.aerobicTable[self.aeTWid.rowCount()-1].append('NI')
                options = IndexedComboBox(self.aeTWid.rowCount()-1, i, self, True)
                options.addItems(self.options)
                self.aeTWid.setCellWidget(self.aeTWid.rowCount()-1, i, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def addRowAnaerobic(self):
        try:
            self.anTWid.setRowCount(self.anTWid.rowCount()+1)
            self.anaerobicTable.append(['Actinobacillus Actinomycetemcomitians'])
            bacteria = IndexedComboBox(self.anTWid.rowCount()-1, 0, self, False)
            bacteria.addItems(self.anaerobicList)
            self.anTWid.setCellWidget(self.anTWid.rowCount()-1, 0, bacteria)
            for i in range(1, self.anTWid.columnCount()):
                self.anaerobicTable[self.anTWid.rowCount()-1].append('NI')
                options = IndexedComboBox(self.anTWid.rowCount()-1, i, self, False)
                options.addItems(self.options)
                self.anTWid.setCellWidget(self.anTWid.rowCount()-1, i, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def delRowAerobic(self):
        if self.aeTWid.rowCount() > 1:
            self.aeTWid.setRowCount(self.aeTWid.rowCount()-1)
            self.aerobicTable.pop()

    def delRowAnaerobic(self):
        if self.anTWid.rowCount() > 1:
            self.anTWid.setRowCount(self.anTWid.rowCount()-1)
            self.anaerobicTable.pop()

    def addColAerobic(self):
        try:
            self.aeTWid.setColumnCount(self.aeTWid.columnCount()+1)
            self.aerobicTable[0].append('Growth')
            header = IndexedComboBox(0, self.aeTWid.columnCount()-1, self, True)
            header.addItems(self.headers)
            self.aeTWid.setCellWidget(0, self.aeTWid.columnCount()-1, header)
            for i in range(1, self.aeTWid.rowCount()):
                self.aerobicTable[i].append('NI')
                options = IndexedComboBox(i, self.aeTWid.columnCount()-1, self, True)
                options.addItems(self.options)
                self.aeTWid.setCellWidget(i, self.aeTWid.columnCount()-1, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def addColAnaerobic(self):
        try:
            self.anTWid.setColumnCount(self.anTWid.columnCount()+1)
            self.anaerobicTable[0].append('Growth')
            header = IndexedComboBox(0, self.anTWid.columnCount()-1, self, False)
            header.addItems(self.headers)
            self.anTWid.setCellWidget(0, self.anTWid.columnCount()-1, header)
            for i in range(1, self.anTWid.rowCount()):
                self.anaerobicTable[i].append('NI')
                options = IndexedComboBox(i, self.anTWid.columnCount()-1, self, False)
                options.addItems(self.options)
                self.anTWid.setCellWidget(i, self.anTWid.columnCount()-1, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def delColAerobic(self):
        if self.aeTWid.columnCount() > 1:
            self.aeTWid.setColumnCount(self.aeTWid.columnCount()-1)
            for row in self.aerobicTable:
                row.pop()

    def delColAnaerobic(self):
        if self.anTWid.columnCount() > 1:
            self.anTWid.setColumnCount(self.anTWid.columnCount()-1)
            for row in self.anaerobicTable:
                row.pop()

    def handleSearchPressed(self):
        try:
            if not self.saID.text().isdigit():
                self.saID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('Cultures', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Collected], [Received], [Reported], [Aerobic Results], [Anaerobic Results], [Comments], [Notes]')
            if self.sample is None:
                self.saID.setText('xxxxxx')
            else:
                self.chID.setText(self.sample[0])
                clinician = self.model.findClinician(self.sample[1])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.patientName.setText(self.sample[2] + " " + self.sample[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.recDate.setDate(self.view.dtToQDate(self.sample[5]))
                self.repDate.setDate(self.view.dtToQDate(self.sample[6]))
                self.aerobicTable = self.resultToTable(self.sample[7])
                self.anaerobicTable = self.resultToTable(self.sample[8])
                self.cText.setText(self.sample[9])
                self.nText.setText(self.sample[10])
                self.initTables()
                self.save.setEnabled(True)
                self.clear.setEnabled(True)
                self.printP.setEnabled(False)
                self.printF.setEnabled(False)
                self.printS.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleSavePressed(self):
        try:
            aerobic = self.tableToResult(self.aerobicTable)
            anaerobic = self.tableToResult(self.anaerobicTable)
            if self.model.addCultureResult(
                int(self.saID.text()),
                self.chID.text(),
                self.view.entries[self.clinDrop.currentText()]['db'],
                self.sample[2],
                self.sample[3],
                self.repDate.date(),
                aerobic,
                anaerobic,
                self.cText.toPlainText(),
                self.nText.toPlainText(),
            ):
                self.handleSearchPressed()
                self.save.setEnabled(False)
                self.clear.setEnabled(False)
                self.printP.setEnabled(True)
                self.printF.setEnabled(True)
                self.printS.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleDirectSmearPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\culture_smear_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge(
                saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
                clinicianName=self.clinDrop.currentText(),
                collected=self.view.fSlashDate(self.sample[4]),
                received=self.view.fSlashDate(self.recDate.date()),
                chartID=self.chID.text(),
                patientName=f'{self.sample[3]}, {self.sample[2]}',
                comments=self.cText.toPlainText(),
                directSmear=self.dText.toPlainText()
            )
            document.write(dst)
            self.view.convertAndPrint(dst)
        except Exception as e:
            self.view.showErrorScreen(e)
    
    def handlePreliminaryPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\culture_prelim_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            clinician=self.clinDrop.currentText().split(', ')
            document.merge(
                saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
                collected=self.view.fSlashDate(self.sample[4]),
                received=self.view.fSlashDate(self.recDate.date()),
                reported=self.view.fSlashDate(self.repDate.date()),
                chartID=self.chID.text(),
                clinicianName=clinician[1] + " " + clinician[0],
                patientName=f'{self.sample[3]}, {self.sample[2]}',
                comments=self.cText.toPlainText(),
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            document.write(dst)
            context = {
                'headers' : ['Aerobic Bacteria']+self.aerobicTable[0][1:],
                'servers': []
            }
            for i in range(1, len(self.aerobicTable)):
                context['servers'].append(self.aerobicTable[i])
            document = DocxTemplate(dst)
            document.render(context)
            document.save(dst)
            self.view.convertAndPrint(dst)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePerioPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\culture_results_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            clinician=self.clinDrop.currentText().split(', ')
            document.merge(
                saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
                collected=self.view.fSlashDate(self.sample[4]),
                received=self.view.fSlashDate(self.recDate.date()),
                reported=self.view.fSlashDate(self.repDate.date()),
                chartID=self.chID.text(),
                clinicianName=clinician[1] + " " + clinician[0],
                patientName=f'{self.sample[3]}, {self.sample[2]}',
                comments=self.cText.toPlainText(),
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            document.write(dst)
            #aerobic
            context = {
                'headers1' : ['Aerobic Bacteria']+self.aerobicTable[0][1:],
                'headers2' : ['Anaerobic Bacteria']+self.anaerobicTable[0][1:],
                'servers1': [],
                'servers2': []
            }
            for i in range(1, len(self.aerobicTable)):
                context['servers1'].append(self.aerobicTable[i])
            for i in range(1, len(self.anaerobicTable)):
                context['servers2'].append(self.anaerobicTable[i])
            document = DocxTemplate(dst)
            document.render(context)
            document.save(dst)
            # document = DocxTemplate(dst)
            # document.render(context)
            # document.save(dst)
            self.view.convertAndPrint(dst)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.saID.clear()
            self.patientName.clear()
            self.clinDrop.setCurrentIndex(0)
            self.chID.clear()
            self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.repDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.cText.clear()
            self.nText.clear()
            self.dText.clear()
            self.initTables()
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleBackPressed(self):
        self.view.showResultEntryNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

class CATResultForm(QMainWindow):
    def __init__(self, model, view):
        super(CATResultForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_CAT_Result_Form.ui", self)
        self.find.setIcon(QIcon('COMBDb/Icon/searchIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItems(self.view.names)

        self.volume.setText("0.00")
        self.collectionTime.setText("0.00")
        self.flowRate.setText("0.00")

        self.volume.editingFinished.connect(lambda: self.lineEdited(True))
        self.collectionTime.editingFinished.connect(lambda: self.lineEdited(False))
        #self.volume.editingFinished.connect(self.volumeLineEdited)
        #self.collectionTime.editingFinished.connect(self.colTimeLineEdited)

        self.save.setEnabled(False)
        self.print.setEnabled(False)
        self.repDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.find.clicked.connect(self.handleSearchPressed)

    def lineEdited(self, arg):
        try:
            lineEdit = self.volume if arg else self.collectionTime
            if float(self.collectionTime.text()) != 0:
                vol = float(self.volume.text())
                colTime = float(self.collectionTime.text())
                value = str(vol if arg else colTime)
                rate = round(vol / colTime, 4)
                lineEdit.setText(value)
                self.flowRate.setText(str(rate)) 
                self.errorMessage.setText(None)
            else:
                self.flowRate.setText("0.00")
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleBackPressed(self):
        self.view.showResultEntryNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleSearchPressed(self):
        try:
            if not self.saID.text().isdigit():
                self.saID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('CATs', int(self.saID.text()), '[Clinician], [First], [Last], [Tech], [Reported], [Type], [Volume (ml)], [Time (min)], [Initial (pH)], [Flow Rate (ml/min)], [Buffering Capacity (pH)], [Strep Mutans (CFU/ml)], [Lactobacillus (CFU/ml)], [Comments], [Notes], [Collected], [Received]')
            if self.sample is None:
                self.saID.setText('xxxxxx')
            else:
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.fName.setText(self.sample[1])
                self.lName.setText(self.sample[2])
                # technician = self.model.tech if self.technician.text() is None else self.model.findTech(self.sample[3], 'Entry, First, Middle, Last, Username, Password, Active')
                #self.technician.setCurrentIndex(self.view.entries['techs'][self.view.fTechName(technician[1], technician[2], technician[3], 'formal')])
                self.repDate.setDate(self.view.dtToQDate(self.sample[4]))
                self.type.setCurrentIndex(self.type.findText(self.sample[5]))
                self.volume.setText(str(self.sample[6]) if self.sample[12] is not None else None)
                self.collectionTime.setText(str(self.sample[7]) if self.sample[12] is not None else None)
                self.initialPH.setText(str(self.sample[8]) if self.sample[12] is not None else None)
                self.flowRate.setText(str(self.sample[9]) if self.sample[12] is not None else None)
                self.bufferingCapacityPH.setText(str(self.sample[10]) if self.sample[12] is not None else None)
                self.strepMutansCount.setText(str(self.sample[11]) if self.sample[12] is not None else None)
                self.lactobacillusCount.setText(str(self.sample[12]) if self.sample[12] is not None else None)
                self.cText.setText(self.sample[13])
                self.nText.setText(self.sample[14])
                self.save.setEnabled(True)
                self.clear.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleSavePressed(self):
        try:
            if float(self.collectionTime.text()) != 0:
                saID = int(self.saID.text())
                #self.saID.setText(str(saID))
                if self.model.addCATResult(
                    saID,
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.fName.text(),
                    self.lName.text(),
                    self.repDate.date(),
                    self.type.currentText(),
                    float(self.volume.text()),
                    float(self.collectionTime.text()),
                    float(self.flowRate.text()),
                    float(self.initialPH.text()),
                    float(self.bufferingCapacityPH.text()),
                    int(self.strepMutansCount.text()),
                    int(self.lactobacillusCount.text()),
                    self.cText.toPlainText(),
                    self.nText.toPlainText()
                ):
                    self.handleSearchPressed()
                    self.save.setEnabled(False)
                    self.clear.setEnabled(False)
                    self.print.setEnabled(True)
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Saved CAT Result Form: " + str(saID))
            else:
                self.flowRate.setText("x.xx")
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Division by zero error")
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.saID.clear()
            self.clinDrop.setCurrentIndex(0)
            self.fName.clear()
            self.lName.clear()
            self.type.setCurrentIndex(0)
            self.volume.setText("0.00")
            self.initialPH.clear()
            self.collectionTime.setText("0.00")
            self.bufferingCapacityPH.clear()
            self.flowRate.setText("0.00")
            self.strepMutansCount.clear()
            self.lactobacillusCount.clear()
            self.repDate.setDate(self.view.dtToQDate(None))
            self.cText.clear()
            self.nText.clear()
            self.save.setEnabled(True)
            self.clear.setEnabled(True)
            self.print.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePrintPressed(self):
        try:
            # print(Path().resolve())
            template = str(Path().resolve())+r'\COMBDb\templates\cat_results_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            clinician=self.clinDrop.currentText().split(', ')
            document.merge(
                saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
                patientName=f'{self.sample[2]}, {self.sample[1]}',
                clinicianName=clinician[1] + " " + clinician[0],
                collected=self.view.fSlashDate(self.sample[13]),
                received=self.view.fSlashDate(self.sample[14]),
                flowRate=str(self.sample[8]),
                bufferingCapacity=str(self.sample[9]),
                smCount='{:.2e}'.format(self.sample[10]),
                lbCount='{:.2e}'.format(self.sample[11]),
                reported=self.view.fSlashDate(self.sample[4]),
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            document.write(dst)
        except Exception as e:
            self.view.showErrorScreen(e)
        try:
            self.view.convertAndPrint(dst)
        except Exception as e:
            self.view.showErrorScreen(e)

class DUWLResultForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLResultForm, self).__init__()
        try:
            self.view = view
            self.model = model
            loadUi("COMBDb/UI Screens/COMBdb_DUWL_Result_Form.ui", self)
            self.find.setIcon(QIcon('COMBDb/Icon/searchIcon.png'))
            self.clinDrop.clear()
            self.clinDrop.addItems(self.view.names)
            self.currentKit = 1
            self.kitList = []
            self.meets = { 'Meets': 1, 'Fails to Meet': 2 }
            self.printList = {}
            self.save.setEnabled(False)
            self.print.setEnabled(False)
            self.repDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.back.clicked.connect(self.handleBackPressed)
            self.home.clicked.connect(self.handleReturnToMainMenuPressed)
            self.save.clicked.connect(self.handleSavePressed)
            self.clear.clicked.connect(self.handleClearPressed)
            self.print.clicked.connect(self.handlePrintPressed)
            self.find.clicked.connect(self.handleSearchPressed)
            self.clearAll.clicked.connect(self.handleClearAllPressed)
            self.remove.clicked.connect(self.handleRemovePressed)
            self.kitTWid.setColumnCount(1)
            self.kitTWid.itemClicked.connect(self.activateRemove)
            self.print.setEnabled(False)
            self.remove.setEnabled(False)
        except Exception as e:
            print(e)

    def activateRemove(self):
        self.remove.setEnabled(True)

    def handleBackPressed(self):
        self.view.showResultEntryNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleSearchPressed(self):
        try:
            if not self.saID.text().isdigit():
                #self.saID.setText('xxxxxx')
                self.handleClearPressed()
                return
            self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Clinician], [Bacterial Count], [CDC/ADA], [Reported], [Comments], [Notes]')
            if self.sample is None:
                #self.saID.setText('xxxxxx')
                self.handleClearPressed()
            else:
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.bacterialCount.setText(str(self.sample[1]) if self.sample[1] else None)
                self.cdcADA.setCurrentIndex(self.meets[self.sample[2]] if self.sample[2] else 0)
                self.repDate.setDate(self.view.dtToQDate(self.sample[3]))
                self.cText.setText(self.sample[4])
                self.nText.setText(self.sample[5])
                self.save.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleSavePressed(self):
        try:
            saID = int(self.saID.text())
            if self.model.addWaterlineResult(
                saID,
                self.view.entries[self.clinDrop.currentText()]['db'],
                self.repDate.date(),
                int(self.bacterialCount.text()),
                self.cdcADA.currentText(),
                self.cText.toPlainText(),
                self.nText.toPlainText()
            ):
                self.kitList.append({
                    'sampleID': f'{str(saID)[0:2]}-{str(saID)[2:]}',
                    'count': self.bacterialCount.text(),
                    'cdcADA': self.cdcADA.currentText()
                })
                self.printList[str(saID)] = self.currentKit-1
                self.currentKit += 1
                self.handleClearPressed()
                self.save.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.saID.setText('xxxxxx')
            self.cText.clear()
            self.nText.clear()
            self.bacterialCount.clear()
            self.cdcADA.setCurrentText(None)
            self.save.setEnabled(True)
            self.clear.setEnabled(True)
            self.clinDrop.setCurrentIndex(0)
            self.updateTable()
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearAllPressed(self):
        try:
            self.kitList.clear()
            self.currentKit = 1
            self.printList.clear()
            self.updateTable()
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleRemovePressed(self):
        try:
            del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
            del self.printList[self.kitTWid.currentItem().text()]
            count = 0
            for key in self.printList.keys():
                self.printList[key] = count
                count += 1
            self.updateTable()
            self.remove.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def updateTable(self):
        try:
            self.kitTWid.setRowCount(len(self.printList.keys()))
            count = 0
            for item in self.printList.keys():
                self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
                count += 1
            if len(self.printList.keys())>0:
                self.print.setEnabled(True)
            else:
                self.print.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePrintPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\duwl_results_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge_rows('sampleID', self.kitList)
            clinician = self.model.findClinician(self.sample[0])
            document.merge(
                reported=self.view.fSlashDate(self.repDate.date()),
                clinicianName=self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3]),
                designation=clinician[3],
                address=clinician[4],
                city=clinician[5],
                state=clinician[6],
                zip=str(clinician[7])
            )
            document.write(dst)
        except Exception as e:
            self.view.showErrorScreen(e)
        try:
            self.view.convertAndPrint(dst)
        except Exception as e:
            self.view.showErrorScreen(e)

class IndexedComboBox(QComboBox):
    def __init__(self, row, column, form, kind):
        super(IndexedComboBox, self).__init__()
        self.row = row
        self.column = column
        self.form = form
        self.kind = kind
        self.currentIndexChanged.connect(self.handleCurrentIndexChanged)

    def handleCurrentIndexChanged(self):
        self.form.updateTable(self.kind, self.row, self.column)
