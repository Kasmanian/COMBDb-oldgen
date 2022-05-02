from __future__ import print_function
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets, QtPrintSupport
from PyQt5.QtWidgets import *
import sys, os, datetime
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
        self.setClinicianList()
        try:
            sys.exit(app.exec())
        except Exception as e:
            print(e)

    def showSetFilePathScreen(self):
        self.setFilePathScreen = SetFilePathScreen(self.model, self)
        self.setFilePathScreen.show()

    def showErrorScreen(self):
        self.setErrorScreen = SetErrorScreen(self.model, self)
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
        except Exception as e:
            print(e)
        try:
            word.ActiveDocument()
        except Exception as e:
            word.Quit()
            print(e)
        os.remove(path)
        try:
            self.showPrintPreview(tempPath)
        except Exception as e:
            print(e)

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

    def setClinicianList(self):
        try:
            self.clinicians = self.model.selectClinicians('Entry, Prefix, First, Last, Designation, Phone, Fax, Email, [Address 1], [Address 2], City, State, Zip, Enrolled, Inactive, Comments')
            self.entries = {}
            self.names = []
            for clinician in self.clinicians:
                name = self.fClinicianName(clinician[1], clinician[2], clinician[3], clinician[4])
                self.names.append(name)
                self.entries[name] = clinician[0]
            self.names.sort()
        except Exception as e:
            print(e)


class SetFilePathScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetFilePathScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Set_File_Path_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()


class SetErrorScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetErrorScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_Error_Window.ui", self)
        # Handle 'OK' button clicked
        self.ok.clicked.connect(self.handleOKPressed)

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
            print(e)

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
            print(e)

    def handleDeactivatePressed(self):
        try:
            if self.selectedTechnician[3] != 'No':
                if self.model.toggleTech(self.selectedTechnician[1], 'No'):
                    self.selectedTechnician[3] = 'No'
                    self.technicianTable.item(self.selectedTechnician[0], 2).setText('No')
        except Exception as e:
            print(e)


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
                self.view.entries[self.clinicianDropDown.currentText()],
                self.firstName.text(),
                self.lastName.text(),
                self.collectionDate.date(),
                self.receivedDate.date(),
                self.comment.toPlainText()
            )
            if sampleID:
                self.sampleID.setText(str(sampleID))
        except Exception as e:
            print(e)
    
    def handlePrintPressed(self):
        template = r'C:\Users\simmsk\Desktop\templates\culture_worksheet_template.docx'
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
            print(e)

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
            print(e)

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
            print(e)
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
                self.view.entries[self.clinicianDropDown.currentText()],
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
            print(e)

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
            template = r'C:\Users\simmsk\Desktop\templates\duwl_label_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            document.merge_rows('sampleID', self.kitList)
        except Exception as e:
            print(e)
        try:
            self.view.convertAndPrint(document, dst)
        except Exception as e:
            print(e)

class DUWLReceiveForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLReceiveForm, self).__init__()
        self.view = view
        self.model = model
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Receive_Form.ui", self)
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

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
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        self.search.clicked.connect(self.handleSearchPressed)
        self.preliminary.clicked.connect(self.handlePreliminaryPressed)

    def handleSearchPressed(self):
        try:
            if not self.sampleID.text().isdigit():
                self.sampleID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('Cultures', int(self.sampleID.text()))
            if self.sample is None:
                self.sampleID.setText('xxxxxx')
            else:
                self.chartNumber.setText(self.sample[0])
                clinician = self.model.findClinician(self.sample[1])
                self.clinician.clear()
                self.clinician.addItem(self.view.fClinicianName(clinician[0], clinician[1], clinician[2], None))
                self.receivedDate.setDate(QDate(self.sample[5].year, self.sample[5].month, self.sample[5].day))
                self.comment.setText(self.sample[6])
            print(self.sample)
        except Exception as e:
            print(e)
    
    def handlePreliminaryPressed(self):
        try:
            template = r'C:\Users\simmsk\Desktop\templates\preliminary_culture_results_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            print(type(self.sample[4]), type(self.receivedDate.date()))
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
            print(e)

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
        self.back.clicked.connect(self.handleBackPressed)
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    def handleBackPressed(self):
        self.view.showResultEntryNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

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