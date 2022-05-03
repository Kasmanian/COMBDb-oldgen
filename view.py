from __future__ import print_function
from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets, QtPrintSupport
from PyQt5.QtWidgets import *
import sys, os, datetime, json
import win32com.client as win32
from mailmerge import MailMerge
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, Qt, QDate

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

    def showGuestLoginScreen(self):
        guestLoginScreen = GuestLoginScreen(self.model, self)
        self.widget.addWidget(guestLoginScreen)
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

    def showEditTechnician(self):
        self.settingsEditTechnician = SettingsEditTechnician(self.model, self)
        self.settingsEditTechnician.show()

    def showSettingsManageArchivesForm(self):
        settingsManageArchivesForm = SettingsManageArchivesForm(self.model, self)
        self.widget.addWidget(settingsManageArchivesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showSettingsManagePrefixesForm(self):
        settingsManagePrefixesForm = SettingsManagePrefixesForm(self.model, self)
        self.widget.addWidget(settingsManagePrefixesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showGuestHomeScreen(self):
        guestHomeScreen = GuestHomeScreen(self.model, self)
        self.widget.addWidget(guestHomeScreen)
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

    def convertAndPrint(self, document, path):
        try:
            document.write(path)
            word = win32.DispatchEx('Word.Application')
            document = word.Documents.Open(path)
            tempPath = path.split('.')[0] + '.html'
            document.SaveAs(tempPath, 10)
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
        if self.model.techLogin(self.usrnm.text(), self.pswd.text()):
            self.view.showAdminHomeScreen()

    def handleGuestLoginPressed(self):
        self.view.showGuestLoginScreen()

class GuestLoginScreen(QMainWindow):
    def __init__(self, model, view):
        super(GuestLoginScreen, self).__init__()
        self.view = view
        self.model = model 
        loadUi("COMBDb/UI Screens/COMBdb_Guest_Login.ui", self)
        self.guestLogin.clicked.connect(self.handleGuestLoginPressed)
    
    def handleGuestLoginPressed(self):
        self.view.showGuestHomeScreen()

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
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.technicianTable.itemSelectionChanged.connect(self.handleTechnicianSelected)
        self.activate.clicked.connect(self.handleActivatePressed)
        self.deactivate.clicked.connect(self.handleDeactivatePressed)
        techs = self.model.selectTechs('Entry, Username, Active')
        self.technicianTable.setRowCount(len(techs)) 
        self.technicianTable.setColumnCount(3)
        try:
            for i in range(0, len(techs)):
                self.technicianTable.setItem(i,0, QTableWidgetItem(str(techs[i][0])))
                self.technicianTable.setItem(i,1, QTableWidgetItem(techs[i][1]))
                self.technicianTable.setItem(i,2, QTableWidgetItem(techs[i][2]))
        except Exception as e:
            self.view.showErrorScreen(e)

    # Method for 'Edit' button functionality
    def handleEditPressed(self):
        self.view.showEditTechnician()

    # Method for 'Back' button functionality
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


class SettingsEditTechnician(QMainWindow):
    # Class for the Edit Technician UI
    def __init__(self, model, view):
        super(SettingsEditTechnician, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Settings_Edit_Technician.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()


class SettingsManageArchivesForm(QMainWindow):
    # Class for the Manage Archives UI
    def __init__(self, model, view):
        super(SettingsManageArchivesForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Settings_Manage_Archives_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    def handleBackPressed(self):
        self.view.showSettingsNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    
class SettingsManagePrefixesForm(QMainWindow):
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
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    def handleBackPressed(self):
        self.view.showSettingsNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()


class GuestHomeScreen(QMainWindow):
    def __init__(self, model, view):
        super(GuestHomeScreen, self).__init__()
        self.view = view
        self.model = model 
        loadUi("COMBDb/UI Screens/COMBdb_Guest_Home_Screen.ui", self)
        self.logout.clicked.connect(self.handleLogoutPressed)

    def handleLogoutPressed(self):
        self.view.showAdminLoginScreen()

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
        self.clinicianDropDown.clear()
        self.clinicianDropDown.addItems(self.view.names)
        self.addClinician.clicked.connect(self.handleAddNewClinicianPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.clear.clicked.connect(self.handleClearPressed)

    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinicianDropDown)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
    
    def handleSavePressed(self):
        try:
            table = 'CATs' if self.cultureTypeDropDown.currentText()=='Caries' else 'Cultures'
            sampleID = self.view.model.addPatientOrder(
                table,
                self.chartNum.text(),
                self.view.entries[self.clinicianDropDown.currentText()]['db'],
                self.firstName.text(),
                self.lastName.text(),
                self.collectionDate.date(),
                self.receivedDate.date(),
                self.comment.toPlainText()
            )
            if sampleID:
                self.sampleID.setText(str(sampleID))
        except Exception as e:
            self.view.showErrorScreen(e)
    
    def handlePrintPressed(self):
        template = str(Path().resolve())+r'\COMBDb\templates\culture_worksheet_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge(
            received=self.receivedDate.date().toString(),
            chartID=self.chartNum.text(),
            clinician=self.clinicianDropDown.currentText(),
            patientName=f'{self.lastName.text()}, {self.firstName.text()}',
            comments=self.comment.toPlainText()
        )
        try:
            self.view.convertAndPrint(document, dst)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.firstName.clear()
            self.lastName.clear()
            self.collectionDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.receivedDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
            self.sampleID.setText('xxxxxx')
            self.chartNum.clear()
            self.comment.clear()
        except Exception as e:
            self.view.showErrorScreen(e)

class AddClinician(QMainWindow):
    def __init__(self, model, view, dropdown):
        super(AddClinician, self).__init__()
        self.view = view
        self.model = model
        self.dropdown = dropdown
        loadUi("COMBDb/UI Screens/COMBdb_Add_New_Clinician.ui", self)
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)

    def handleSavePressed(self):
        try:
            self.model.addClinician(
                self.title.currentText(),
                self.firstName.text(),
                self.lastName.text(),
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
                self.comment.toPlainText()
            )
            self.view.setClinicianList()
            self.dropdown.clear()
            self.dropdown.addItems(self.view.names)
        except Exception as e:
            self.view.showErrorScreen(e)
        finally:
            self.close()

    def handleBackPressed(self):
        self.close()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

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
        self.currentKit = 1
        self.kitList = []
        self.kitNumber.setText('1')
        self.clinicianDropDown.clear()
        self.clinicianDropDown.addItems(self.view.names)
        self.next.setEnabled(False)
        self.print.setEnabled(False)
        self.shippingDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.addClinician.clicked.connect(self.handleAddClinicianPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.next.clicked.connect(self.handleNextPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.clearAll.clicked.connect(self.handleClearAllPressed)
        self.print.clicked.connect(self.handlePrintPressed)


    def handleAddClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinicianDropDown)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleSavePressed(self):
        try:
            sampleID = self.view.model.addWaterlineOrder(
                self.view.entries[self.clinicianDropDown.currentText()]['db'],
                self.shippingDate.date(),
                self.comment.toPlainText()
            )
            if sampleID:
                self.sampleID.setText(str(sampleID))
                self.save.setEnabled(False)
                self.next.setEnabled(True)
                self.clear.setEnabled(False)
                self.print.setEnabled(True)
                self.kitList.append({
                    'sampleID': f'{str(sampleID)[0:2]}-{str(sampleID)[2:]}',
                    'operatory': 'Operatory___________________________',
                    'collected': 'Collection Date______________________',
                    'clngagent': 'Cleaning Agent______________________'
                })
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleNextPressed(self):
        self.currentKit += 1
        self.handleClearPressed()
        self.save.setEnabled(True)
        self.next.setEnabled(False)
        self.clear.setEnabled(True)
        self.print.setEnabled(False)

    def handleClearPressed(self):
        self.kitNumber.setText(str(self.currentKit))
        self.sampleID.setText('xxxxxx')
        self.comment.clear()
        self.save.setEnabled(True)
        self.clear.setEnabled(True)

    def handleClearAllPressed(self):
        self.kitList.clear()
        self.currentKit = 1
        self.handleClearPressed()

    def handlePrintPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\duwl_label_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge_rows('sampleID', self.kitList)
        except Exception as e:
            self.view.showErrorScreen(e)
        try:
            self.view.convertAndPrint(document, dst)
        except Exception as e:
            self.view.showErrorScreen(e)

class DUWLReceiveForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLReceiveForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Receive_Form.ui", self)
        #self.currentKit = 1
        self.kitList = []
        #self.kitNumber.setText('1')
        self.clinicianDropDown.clear()
        self.clinicianDropDown.addItems(self.view.names)
        #self.next.setEnabled(False)
        self.save.setEnabled(False)
        self.print.setEnabled(False)
        self.receivedDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.collectedDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        #self.next.clicked.connect(self.handleNextPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        #self.clearAll.clicked.connect(self.handleClearAllPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.search.clicked.connect(self.handleSearchPressed)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleSearchPressed(self):
        try:
            if not self.sampleNum_2.text().isdigit():
                self.sampleNum_2.setText('xxxxxx')
                return
            self.sample = self.model.findSample('Waterlines', int(self.sampleNum_2.text()), 'Clinician, Comments, OperatoryID, Product, Procedure, Collected, Received')
            if self.sample is None:
                self.sampleNum_2.setText('xxxxxx')
            else:
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinicianDropDown.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.comment.setText(self.sample[1])
                self.operatory.setText(self.sample[2])
                self.product.setText(self.sample[3])
                self.procedure.setText(self.sample[4])
                self.collectedDate.setDate(self.view.dtToQDate(self.sample[5]))
                self.receivedDate.setDate(self.view.dtToQDate(self.sample[6]))
                self.save.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleSavePressed(self):
        try:
            sampleID = int(self.sampleNum_2.text())
            #self.sampleNum_2.setText(str(sampleID))
            if self.model.addWaterlineReceiving(
                sampleID,
                self.operatory.text(),
                self.view.entries[self.clinicianDropDown.currentText()]['db'],
                self.collectedDate.date(),
                self.receivedDate.date(),
                self.product.text(),
                self.procedure.text(),
                self.comment.toPlainText()
            ):
                self.save.setEnabled(False)
                #self.next.setEnabled(True)
                self.clear.setEnabled(False)
                self.print.setEnabled(True)
                # self.kitList.append({
                #     'sampleID': f'{str(sampleID)[0:2]}-{str(sampleID)[2:]}',
                #     'operatory': 'Operatory___________________________',
                #     'collected': 'Collection Date______________________',
                #     'clngagent': 'Cleaning Agent______________________'
                # })
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleNextPressed(self):
        self.currentKit += 1
        self.handleClearPressed()
        self.save.setEnabled(True)
        self.next.setEnabled(False)
        self.clear.setEnabled(True)
        self.print.setEnabled(False)

    def handleClearPressed(self):
        # self.kitNumber.setText(str(self.currentKit))
        self.sampleNum_2.setText('xxxxxx')
        self.comment.clear()
        self.save.setEnabled(True)
        self.clear.setEnabled(True)

    def handleClearAllPressed(self):
        self.kitList.clear()
        self.currentKit = 1
        self.handleClearPressed()

    def handlePrintPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\duwl_label_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge_rows('sampleID', self.kitList)
        except Exception as e:
            self.view.showErrorScreen(e)
        try:
            self.view.convertAndPrint(document, dst)
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
        self.clinician.clear()
        self.clinician.addItems(self.view.names)
        self.receivedDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.dateReported.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.search.clicked.connect(self.handleSearchPressed)
        self.preliminary.clicked.connect(self.handlePreliminaryPressed)
        self.tableWidget.itemSelectionChanged.connect(self.handleCellChanged)
        #testbox = QComboBox()
        #self.tableWidget.setCellWidget(0, 0, testbox)
        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.setColumnCount(0)
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
                self.optionIndexes = { 'NI': 0, 'L': 1, 'M': 2, 'H': 3, 'P': 4, 'N': 5 }
                for antibiotics in self.antibioticPrefixes.keys():
                    self.headers.append(antibiotics)
                    self.headerIndexes[antibiotics] = len(self.headers)-1
            self.addRow.clicked.connect(self.addRowAerobic)
            self.addRow_2.clicked.connect(self.addRowAnaerobic)
            self.removeRow.clicked.connect(self.delRowAerobic)
            self.removeRow_2.clicked.connect(self.delRowAnaerobic)
            self.addColumn.clicked.connect(self.addColAerobic)
            self.addColumn_2.clicked.connect(self.addColAnaerobic)
            self.removeColumn.clicked.connect(self.delColAerobic)
            self.removeColumn_2.clicked.connect(self.delColAnaerobic)
            self.aerobicTable = self.resultToTable(None)
            self.anaerobicTable = self.resultToTable(None)
            self.initTables()
        except Exception as e:
            self.view.showErrorScreen(e)

    # def initAerobicTable(self):
    #     try:
    #         currentHeaders1 = ['Growth', 'B-lac', 'PEN', 'AMP', 'CC', 'TET', 'CEP', 'ERY']
    #         currentHeaders2 = ['Growth', 'B-lac', 'PEN', 'AMP', 'CC', 'TET', 'CEP', 'ERY', 'MET']
    #         if self.aerobicResults is not None:
    #             if len(self.aerobicResults)>0:
    #                 currentHeaders1.clear()
    #                 preheader1 = self.aerobicResults[0]['result']
    #                 for header in preheader1:
    #                     currentHeaders1.append(header.split('=')[0])
    #         if self.anaerobicResults is not None:
    #             if len(self.anaerobicResults)>0:
    #                 currentHeaders2.clear()
    #                 preheader2 = self.anaerobicResults[0]['result']
    #                 for header in preheader2:
    #                     currentHeaders2.append(header.split('=')[0])
    #         self.tableWidget.setRowCount(0)
    #         self.tableWidget.setRowCount(1)
    #         self.tableWidget_2.setRowCount(0)
    #         self.tableWidget_2.setRowCount(1)
    #         self.tableWidget.setColumnCount(0)
    #         self.tableWidget.setColumnCount(9)
    #         self.tableWidget_2.setColumnCount(0)
    #         self.tableWidget_2.setColumnCount(10)
    #         self.tableWidget.setItem(0,0, QTableWidgetItem('Bacteria'))
    #         self.tableWidget_2.setItem(0,0, QTableWidgetItem('Bacteria'))
    #         count = 1
    #         for header in currentHeaders1:
    #             column = QComboBox()
    #             column.addItems(self.headers)
    #             column.setCurrentIndex(self.headerIndexes[header])
    #             self.tableWidget.setCellWidget(0, count, column)
    #             count += 1
    #         count = 1
    #         for header in currentHeaders2:
    #             column = QComboBox()
    #             column.addItems(self.headers)
    #             column.setCurrentIndex(self.headerIndexes[header])
    #             self.tableWidget_2.setCellWidget(0, count, column)
    #             count += 1
    #     except Exception as e:
    #         self.view.showErrorScreen(e)

    def handleCellChanged(self):
        self.view.showErrorScreen('updated!')

    def initTables(self):
        try:
            self.tableWidget.setRowCount(0)
            self.tableWidget.setRowCount(len(self.aerobicTable))
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.setRowCount(len(self.anaerobicTable))
            self.tableWidget.setColumnCount(0)
            self.tableWidget.setColumnCount(len(self.aerobicTable[0]))
            self.tableWidget_2.setColumnCount(0)
            self.tableWidget_2.setColumnCount(len(self.anaerobicTable[0]))
            #aerobic
            self.tableWidget.setItem(0,0, QTableWidgetItem('Bacteria'))
            for i in range(1, len(self.aerobicTable[0])):
                for j in range(0, len(self.aerobicTable)):
                    item = IndexedComboBox(j, i)
                    if j>0:
                        item.addItems(self.aerobicList)
                        item.setCurrentIndex(self.optionIndexes[self.aerobicTable[j][i]])
                    else:
                        item.addItems(self.headers)
                        item.setCurrentIndex(self.headerIndexes[self.aerobicTable[j][i]])
                    self.tableWidget.setCellWidget(j, i, item)
            #anaerobic
            self.tableWidget_2.setItem(0,0, QTableWidgetItem('Bacteria'))
            for i in range(1, len(self.anaerobicTable[0])):
                for j in range(0, len(self.anaerobicTable)):
                    item = IndexedComboBox(j, i)
                    if j>0:
                        item.addItems(self.anaerobicList)
                        item.setCurrentIndex(self.optionIndexes[self.anaerobicTable[j][i]])
                    else:
                        item.addItems(self.headers)
                        item.setCurrentIndex(self.headerIndexes[self.anaerobicTable[j][i]])
                    self.tableWidget_2.setCellWidget(j, i, item)
        except Exception as e:
            self.view.showErrorScreen(e)

    def resultToTable(self, result):
        if result is not None:
            result = result.spilt('/')
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
            self.tableWidget.setRowCount(self.tableWidget.rowCount()+1)
            bacteria = QComboBox()
            bacteria.addItems(self.aerobicList)
            self.tableWidget.setCellWidget(self.tableWidget.rowCount()-1, 0, bacteria)
            for i in range(1, self.tableWidget.columnCount()):
                options = QComboBox()
                options.addItems(list(self.growthList) + list(self.blacList))
                self.tableWidget.setCellWidget(self.tableWidget.rowCount()-1, i, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def addRowAnaerobic(self):
        try:
            self.tableWidget_2.setRowCount(self.tableWidget_2.rowCount()+1)
            bacteria = QComboBox()
            bacteria.addItems(self.anaerobicList)
            self.tableWidget_2.setCellWidget(self.tableWidget_2.rowCount()-1, 0, bacteria)
            for i in range(1, self.tableWidget_2.columnCount()):
                options = QComboBox()
                options.addItems(list(self.growthList) + list(self.blacList))
                self.tableWidget_2.setCellWidget(self.tableWidget_2.rowCount()-1, i, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def delRowAerobic(self):
        if self.tableWidget.rowCount() > 1:
            self.tableWidget.setRowCount(self.tableWidget.rowCount()-1)

    def delRowAnaerobic(self):
        if self.tableWidget_2.rowCount() > 1:
            self.tableWidget_2.setRowCount(self.tableWidget_2.rowCount()-1)

    def addColAerobic(self):
        try:
            self.tableWidget.setColumnCount(self.tableWidget.columnCount()+1)
            header = QComboBox()
            header.addItems(self.headers)
            self.tableWidget.setCellWidget(0, self.tableWidget.columnCount()-1, header)
            for i in range(1, self.tableWidget.rowCount()):
                options = QComboBox()
                options.addItems(list(self.growthList) + list(self.blacList))
                self.tableWidget.setCellWidget(i, self.tableWidget.columnCount()-1, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def addColAnaerobic(self):
        try:
            self.tableWidget_2.setColumnCount(self.tableWidget_2.columnCount()+1)
            header = QComboBox()
            header.addItems(self.headers)
            self.tableWidget_2.setCellWidget(0, self.tableWidget_2.columnCount()-1, header)
            for i in range(1, self.tableWidget_2.rowCount()):
                options = QComboBox()
                options.addItems(list(self.growthList) + list(self.blacList))
                self.tableWidget_2.setCellWidget(i, self.tableWidget_2.columnCount()-1, options)
        except Exception as e:
            self.view.showErrorScreen(e)

    def delColAerobic(self):
        if self.tableWidget.columnCount() > 1:
            self.tableWidget.setColumnCount(self.tableWidget.columnCount()-1)

    def delColAnaerobic(self):
        if self.tableWidget_2.columnCount() > 1:
            self.tableWidget_2.setColumnCount(self.tableWidget_2.columnCount()-1)

    def handleSearchPressed(self):
        try:
            if not self.sampleID.text().isdigit():
                self.sampleID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('Cultures', int(self.sampleID.text()), '[ChartID], [Clinician], [First], [Last], [Collected], [Received], [Reported], [Aerobic Results], [Anaerobic Results], [Comments]')
            if self.sample is None:
                self.sampleID.setText('xxxxxx')
            else:
                self.chartNumber.setText(self.sample[0])
                clinician = self.model.findClinician(self.sample[1])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinician.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.receivedDate.setDate(self.view.dtToQDate(self.sample[5]))
                self.dateReported.setDate(self.view.dtToQDate(self.sample[6]))
                self.results = {}
                self.results['aerobic'] = {}
                self.comment.setText(self.sample[6])
        except Exception as e:
            self.view.showErrorScreen(e)
    
    def handlePreliminaryPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\preliminary_culture_results_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge(
                sampleID=f'{self.sampleID.text()[0:2]}-{self.sampleID.text()[2:6]}',
                collected=self.view.fSlashDate(self.sample[4]),
                received=self.view.fSlashDate(self.receivedDate.date()),
                reported=self.view.fSlashDate(self.dateReported.date()),
                chartID=self.chartNumber.text(),
                clinicianName=self.clinician.currentText(),
                patientName=f'{self.sample[3]}, {self.sample[2]}',
                comments=self.comment.toPlainText(),
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            self.view.convertAndPrint(document, dst)
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
        self.clinicianDropDown.clear()
        self.clinicianDropDown.addItems(self.view.names)
        self.save.setEnabled(False)
        self.print.setEnabled(False)
        self.dateReported.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.search.clicked.connect(self.handleSearchPressed)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    def handleSearchPressed(self):
        try:
            if not self.sampleID.text().isdigit():
                self.sampleID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('CATs', int(self.sampleID.text()), '[Clinician], [First], [Last], [Tech], [Reported], [Volume (ml)], [Time (min)], [Initial (pH)], [Flow Rate (ml/min)], [Buffering Capacity (pH)], [Strep Mutans (CFU/ml)], [Lactobacillus (CFU/ml)], [Comments], [Collected], [Received]')
            if self.sample is None:
                self.sampleID.setText('xxxxxx')
            else:
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinicianDropDown.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.firstName.setText(self.sample[1])
                self.lastName.setText(self.sample[2])
                technician = self.model.tech if self.technician.text() is None else self.model.findTech(self.sample[3], 'Entry, First, Middle, Last, Username, Password, Active')
                #self.technician.setCurrentIndex(self.view.entries['techs'][self.view.fTechName(technician[1], technician[2], technician[3], 'formal')])
                self.dateReported.setDate(self.view.dtToQDate(self.sample[4]))
                self.volume.setText(str(self.sample[5]) if self.sample[11] is not None else None)
                self.collectionTime.setText(str(self.sample[6]) if self.sample[11] is not None else None)
                self.initialPH.setText(str(self.sample[7]) if self.sample[11] is not None else None)
                self.flowRate.setText(str(self.sample[8]) if self.sample[11] is not None else None)
                self.bufferingCapacityPH.setText(str(self.sample[9]) if self.sample[11] is not None else None)
                self.strepMutansCount.setText(str(self.sample[10]) if self.sample[11] is not None else None)
                self.lactobacillusCount.setText(str(self.sample[11]) if self.sample[11] is not None else None)
                self.comment.setText(self.sample[12])
                self.save.setEnabled(True)
                self.clear.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleSavePressed(self):
        try:
            sampleID = int(self.sampleID.text())
            #self.sampleNum_2.setText(str(sampleID))
            if self.model.addCATResult(
                sampleID,
                self.view.entries[self.clinicianDropDown.currentText()]['db'],
                self.firstName.text(),
                self.lastName.text(),
                self.dateReported.date(),
                float(self.volume.text()),
                float(self.collectionTime.text()),
                float(self.flowRate.text()),
                float(self.initialPH.text()),
                float(self.bufferingCapacityPH.text()),
                int(self.strepMutansCount.text()),
                int(self.lactobacillusCount.text()),
                self.comment.toPlainText(),
            ):
                self.handleSearchPressed()
                self.save.setEnabled(False)
                self.clear.setEnabled(False)
                self.print.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleClearPressed(self):
        try:
            self.sampleID.setText('xxxxxx')
            self.clinicianDropDown.setCurrentIndex(0)
            self.firstName.clear()
            self.lastName.clear()
            self.volume.clear()
            self.initialPH.clear()
            self.collectionTime.clear()
            self.bufferingCapacityPH.clear()
            self.flowRate.clear()
            self.strepMutansCount.clear()
            self.technician.clear()
            self.lactobacillusCount.clear()
            self.dateReported.setDate(self.view.dtToQDate(None))
            self.comment.clear()
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
            document.merge(
                sampleID=f'{self.sampleID.text()[0:2]}-{self.sampleID.text()[2:6]}',
                patientName=f'{self.sample[2]}, {self.sample[1]}',
                clinicianName=self.clinicianDropDown.currentText(),
                collected=self.view.fSlashDate(self.sample[13]),
                received=self.view.fSlashDate(self.sample[14]),
                flowRate=str(self.sample[8]),
                bufferingCapacity=str(self.sample[9]),
                smCount='{:.2e}'.format(self.sample[10]),
                lbCount='{:.2e}'.format(self.sample[11]),
                reported=self.view.fSlashDate(self.sample[4]),
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
        except Exception as e:
            self.view.showErrorScreen(e)
        try:
            self.view.convertAndPrint(document, dst)
        except Exception as e:
            self.view.showErrorScreen(e)

class DUWLResultForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLResultForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Result_Form.ui", self)
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    def handleBackPressed(self):
        self.view.showResultEntryNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

class IndexedComboBox(QComboBox):
    def __init__(self, row, column):
        super(IndexedComboBox, self).__init__()
        self.row = row
        self.column = column