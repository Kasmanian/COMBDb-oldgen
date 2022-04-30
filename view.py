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
        # Launch application straight to Admin Login Screen
        screen = AdminLoginScreen(model, self)
        self.widget = QtWidgets.QStackedWidget()
        self.widget.addWidget(screen)
        # self.widget.setFixedHeight(1200)
        # self.widget.setFixedWidth(1600)
        # self.widget.setWindowTitle("Login Screen")
        self.widget.setGeometry(10,10,1000,800)
        self.widget.showMaximized()
        #self.widget.show()
        try:
            sys.exit(app.exec())
        except Exception as e:
            print(e)

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

    def showSettingsScreen(self):
        settingsScreen = SettingsScreen(self.model, self)
        self.widget.addWidget(settingsScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showGuestHomeScreen(self):
        guestHomeScreen = GuestHomeScreen(self.model, self)
        self.widget.addWidget(guestHomeScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCultureOrderNav(self): # This is the Nav Menu for Culture Orders and below are subsequent screens
        self.cultureOrderNav = CultureOrderNav(self.model, self)
        self.cultureOrderNav.show()
        # self.widget.addWidget(cultureOrderNav)
        # self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCultureOrderForm(self):
        cultureOrderForm = CultureOrderForm(self.model, self)
        self.widget.addWidget(cultureOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAddClinicianScreen(self, dropdown):
        self.addClinician = AddClinician(self.model, self, dropdown)
        self.addClinician.show()
        #self.widget.addWidget(addClinician)
        #self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    #def showCATOrderForm(self):
        #catOrderForm = CATOrderForm(self.model, self)
        #self.widget.addWidget(catOrderForm)
        #self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLNav(self): # This is a sub Nav Menu for DUWL Order Culture or Receive Culture
        self.duwlNav = DUWLNav(self.model, self)
        self.duwlNav.show()
        # self.widget.addWidget(duwlNav)
        # self.widget.show()
        # self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLOrderForm(self):
        duwlOrderForm = DUWLOrderForm(self.model, self)
        self.widget.addWidget(duwlOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLReceiveForm(self):
        duwlReceiveForm = DUWLReceiveForm(self.model, self)
        self.widget.addWidget(duwlReceiveForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showResultEntryNav(self): # This is the Nav Menu for Result Entry and below are subsequent screens
        self.resultEntryNav = ResultEntryNav(self.model, self)
        self.resultEntryNav.show()
        # self.widget.addWidget(resultEntryNav)
        # self.widget.setCurrentIndex(self.widget.currentIndex()+1)

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
        document.write(path)
        word = win32.gencache.EnsureDispatch('Word.Application')
        document = word.Documents.Open(path)
        tempPath = path.split('.')[0] + '.html'
        document.SaveAs(tempPath, 10)
        document.Close()
        try:
            word.ActiveDocument()
        except Exception:
            word.Quit()
        os.remove(path)
        try:
            self.showPrintPreview(tempPath)
        except Exception as e:
            print(e)
        pass

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

class AdminLoginScreen(QMainWindow):
    # Class for the Login Screen UI
    def __init__(self, model, view):
        super(AdminLoginScreen, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Login Screen
        loadUi("COMBDb/UI Screens/COMBdb_Admin_Login.ui", self)
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        # Handle 'Login' button clicked
        self.login.clicked.connect(self.handleLoginPressed)
        # self.guestLogin.clicked.connect(self.handleGuestLoginPressed)

    def handleLoginPressed(self):
        # If credential check is successful, display Admin Home Screen
        if self.model.techLogin(self.usrnm.text(), self.pswd.text()):
            self.view.showAdminHomeScreen()
            return

    # Method for 'Sign in as guest' button functionality
    def handleGuestLoginPressed(self):
        self.view.showGuestLoginScreen()


class GuestLoginScreen(QMainWindow):
    # Class for the Guest Login Screen UI
    def __init__(self, model, view):
        super(GuestLoginScreen, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Guest Login Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Guest_Login.ui", self)
        # Handle 'Login' button clicked
        self.guestLogin.clicked.connect(self.handleGuestLoginPressed)
    
    def handleGuestLoginPressed(self):
        self.view.showGuestHomeScreen()


class AdminHomeScreen(QMainWindow):
    # Class for the Admin Home Screen UI
    def __init__(self, model, view):
        super(AdminHomeScreen, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Admin_Home_Screen.ui", self)
        # Handle 'Culture Order Forms' button clicked
        self.cultureOrder.clicked.connect(self.handleCultureOrderFormsPressed)
        # Handle 'Result Entry' button clicked
        self.resultEntry.clicked.connect(self.handleResultEntryPressed)
        # Handle 'Settings' button clicked
        self.settings.clicked.connect(self.handleSettingsPressed)
        # Handle 'Logout' button clicked
        self.logout.clicked.connect(self.handleLogoutPressed)

    # Method for 'Culture Order Forms' button functionality
    def handleCultureOrderFormsPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Result Entry' button functionality
    def handleResultEntryPressed(self):
        self.view.showResultEntryNav()

    # Method for 'Setting' button functionality
    def handleSettingsPressed(self):
        self.view.showSettingsScreen()

    # Method for 'Logout' button functionality
    def handleLogoutPressed(self):
        self.view.showAdminLoginScreen()


class SettingsScreen(QMainWindow):
     # Class for the Settings Screen UI
    def __init__(self, model, view):
        super(SettingsScreen, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Admin_Settings_Screen.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()


class GuestHomeScreen(QMainWindow):
    # Class for the Guest Home Screen UI
    def __init__(self, model, view):
        super(GuestHomeScreen, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Guest_Home_Screen.ui", self)
        # Handle 'Logout' button clicked
        self.logout.clicked.connect(self.handleLogoutPressed)

    # Method for 'Culture' button functionality
    def handleLogoutPressed(self):
        self.view.showAdminLoginScreen()


class CultureOrderNav(QMainWindow):
    # Class for the Culture Order Forms Navigation Screen UI
    def __init__(self, model, view):
        super(CultureOrderNav, self).__init__()
        self.view = view
        self.model = model
        #self.setFixedHeight(328)
        #self.setFixedWidth(634)
        # Load the .ui file of the Culture Order Navigation Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Forms_Nav.ui", self)
        # Handle 'Culture' button clicked
        self.culture.clicked.connect(self.handleCulturePressed)
        # Handle 'CAT' button clicked
        # self.cat.clicked.connect(self.handleCATPressed)
        # Handle 'DUWL' button clicked
        self.duwl.clicked.connect(self.handleDUWLPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Culture' button functionality
    def handleCulturePressed(self):
        self.view.showCultureOrderForm()
        self.close()

    # Method for 'CAT' button functionality
    # def handleCATPressed(self):
        # self.view.showCATOrderForm()

    # Method for 'DUWL' button functionality
    def handleDUWLPressed(self):
        # self.view.showDUWLOrderForm()
        self.view.showDUWLNav()
        self.close()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()
        #self.view.showPreviousScreen()


class CultureOrderForm(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(CultureOrderForm, self).__init__()
        self.view = view
        self.model = model
        self.clinicians = model.selectClinicians()
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Form.ui", self)
        self.entries = {}
        try:
            c = []
            for clinician in self.clinicians:
                entry = clinician[0]
                first = clinician[1] if clinician[1] is not None else ''
                last = clinician[2]+', ' if clinician[2] is not None else ''
                name = last+first if  clinician[3] is None else clinician[3]
                c.append(name)
                self.entries[name] = entry
            c.sort()
            self.clinicianDropDown.clear()
            self.clinicianDropDown.addItems(c)
        except Exception as e:
            print(e)
        # Handle 'Add New Clinician' button clicked
        self.addClinician.clicked.connect(self.handleAddNewClinicianPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        # Handle 'Save' button clicked
        self.save.clicked.connect(self.handleSavePressed)

        self.print.clicked.connect(self.handlePrintPressed)

    # Method for 'Add New Clinicians functionality
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinicianDropDown)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        #self.view.showPreviousScreen()
    
    # Method for entering new Patient Sample Data
    def handleSavePressed(self):
        try:
            table = 'CATs' if self.cultureTypeDropDown.currentText()=='Caries' else 'Cultures'
            sampleID = self.view.model.addPatientOrder(
                table,
                self.chartNum.text(),
                self.entries[self.clinicianDropDown.currentText()],
                self.firstName.text(),
                self.lastName.text(),
                self.collectionDat.date(),
                self.receivedDate.date(),
                self.comment.toPlainText()
            )
            if sampleID:
                self.sampleNum_2.setText(str(sampleID))
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

class AddClinician(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view, dropdown):
        super(AddClinician, self).__init__()
        self.view = view
        self.model = model
        self.dropdown = dropdown
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Add_New_Clinician.ui", self)
        # Handle 'Add New Clinician' button clicked
        #self.save.clicked.connect(self.handleSavePressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

        self.save.clicked.connect(self.handleSavePressed)

    # Method for 'Add New Clinicians functionality
    def handleSavePressed(self):
        try:
            clinician = self.view.fClinicianName(
                self.title.text(),
                self.firstName.text(),
                self.lastName.text(),
                self.designation.text()
            )
            self.dropdown.addItem(clinician)
            self.model.addClinician(
                self.title.text(),
                self.firstName.text(),
                self.lastName.text(),
                self.designation.text(),
                self.phone.text(),
                self.fax.text(),
                self.email.text(),
                self.address1.text(),
                self.address2.text(),
                self.city.text(),
                self.state.text(),
                self.zip.text(),
                None,
                None,
                self.comment.toPlainText()
            )
        except Exception as e:
            print(e)
        finally:
            self.close()
        # Save the form and add Clinician to database

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()
        #self.view.showPreviousScreen()


""" class CATOrderForm(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(CATOrderForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_CAT_Order_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
"""

class DUWLNav(QMainWindow):
    def __init__(self, model, view):
        super(DUWLNav, self).__init__()
        self.view = view
        self.model = model
        #self.setFixedHeight(250)
        #self.setFixedWidth(600)
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Nav.ui", self)
        # Handle 'Order Culture' button clicked
        self.orderCulture.clicked.connect(self.handleOrderCulturePressed)

        # Handle 'Receiving Culture' button clicked
        self.receivingCulture.clicked.connect(self.handleReceivingCulturePressed)

        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Order Culture' button functionality
    def handleOrderCulturePressed(self):
        self.close()
        self.view.showDUWLOrderForm()

    # Method for 'Receiving Culture' button functionality
    def handleReceivingCulturePressed(self):
        self.close()
        self.view.showDUWLReceiveForm()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()
        self.view.showCultureOrderNav()


class DUWLOrderForm(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(DUWLOrderForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Order_Form.ui", self)
        # Handle 'Add New Clinician' button clicked
        self.addClinician.clicked.connect(self.handleAddClinicianPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Add New Clinician' button functionality
    def handleAddClinicianPressed(self):
        self.view.showAddClinicianScreen()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()


class DUWLReceiveForm(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(DUWLReceiveForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Receive_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()


class ResultEntryNav(QMainWindow):
    # Class for the Result Entry Nav UI
    def __init__(self, model, view):
        super(ResultEntryNav, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Result_Entry_Forms_Nav.ui", self)
        # Handle 'Culture' button clicked
        self.culture.clicked.connect(self.handleCulturePressed)
        # Handle 'CAT' button clicked
        self.cat.clicked.connect(self.handleCATPressed)
        # Handle 'DUWL' button clicked
        self.duwl.clicked.connect(self.handleDUWLPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Culture' button functionality
    def handleCulturePressed(self):
        self.close()
        self.view.showCultureResultForm()

    # Method for 'CAT' button functionality
    def handleCATPressed(self):
        self.close()
        self.view.showCATResultForm()

    # Method for 'DUWL' button functionality
    def handleDUWLPressed(self):
        self.close()
        self.view.showDUWLResultForm()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()


class CultureResultForm(QMainWindow):
    # Class for the Culture Result Form UI
    def __init__(self, model, view):
        super(CultureResultForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Result Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Result_Form.ui", self)
        # Handle 'Search button clicked
        #self.search.clicked.connect(self.handleSearchPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        # Handle 'Search' button clicked
        self.search.clicked.connect(self.handleSearchPressed)
        # Hanlde 'Preliminary' button clicked
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

    # Method for 'Search' button functionality
    #def handleSearchPressed(self):
        #do stuff
        #print("You searched")

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        #self.view.showPreviousScreen()

class CATResultForm(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(CATResultForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_CAT_Result_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()


class DUWLResultForm(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(DUWLResultForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_DUWL_Result_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()