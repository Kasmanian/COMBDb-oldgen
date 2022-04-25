from __future__ import print_function
from tkinter import CENTER, Button
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets, QtPrintSupport, QtCore
from PyQt5.QtWidgets import *
import sys, shutil, os, webbrowser
import win32com.client as win32
from mailmerge import MailMerge
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl, Qt

def passPrintPrompt(boolean):
        print('Done printing')

class View:
    def __init__(self, model):
        self.model = model
        app = QApplication(sys.argv)
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
        except:
            print("Exiting")

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

    def showAddClinicianScreen(self):
        addClinician = AddClinician(self.model, self)
        self.widget.addWidget(addClinician)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

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
        resultEntryNav = ResultEntryNav(self.model, self)
        self.widget.addWidget(resultEntryNav)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

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

    def printFromTemplate(self, template, fieldData):
        dst = template.split('\\')
        dst[len(dst)-1] = 'temp.docx'
        dst = '\\'.join(dst)
        shutil.copyfile(template, dst)
        document = MailMerge(dst)

        pass

    def showPrintNav(self, path):
        self.printNav = PrintNav(path)
        self.printNav.show()

    def showPrintPreview(self, path):
        self.web = QWebEngineView()
        self.web.setContextMenuPolicy(Qt.ActionsContextMenu)
        printAction = QAction('Print', self.web)
        printAction.triggered.connect(self.showPrintPrompt)
        self.web.addAction(printAction)
        self.web.load(QUrl.fromLocalFile(path))
        self.web.showMaximized()
        # self.web.page().windowCloseRequested.connect(self.showPrintPrompt)

    def showPrintPrompt(self):
        print('Print prompt entered')
        self.dialog = QtPrintSupport.QPrintDialog()
        if self.dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.web.page().print(self.dialog.printer(), passPrintPrompt)
            print('Clicked print!')

    #def showPreviousScreen(self):
        #elf.widget.setCurrentIndex(self.widget.currentIndex()-1)


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
        self.guestLogin.clicked.connect(self.handleGuestLoginPressed)

    def handleLoginPressed(self):
        # If credential check is successful, display Admin Home Screen
        if self.model.techLogin(self.usrnm.text(), self.pswd.text()):
            print('Success! Logging you in...')
            self.view.showAdminHomeScreen()
            return
        print('Wrong username or password')

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
        print('Logged out...')
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
        print('Logged out...')
        self.view.showAdminLoginScreen()


class CultureOrderNav(QMainWindow):
    # Class for the Culture Order Forms Navigation Screen UI
    def __init__(self, model, view):
        super(CultureOrderNav, self).__init__()
        self.view = view
        self.model = model
        self.setFixedHeight(328)
        self.setFixedWidth(634)
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
        self.view.showAddClinicianScreen()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        #self.view.showPreviousScreen()
    
    # Method for entering new Patient Sample Data
    def handleSavePressed(self):
        table = 'CATs' if self.cultureTypeDropDown.currentText()=='Caries' else 'Cultures'
        sampleID = self.view.model.addPatientOrder(
            table,
            self.chartNum.text(),
            self.entries[self.clinicianDropDown.currentText()],
            self.firstName.text(),
            self.lastName.text(),
            self.collectionDate.date(),
            self.receivedDate.date(),
            self.comment.toPlainText()
        )
        if sampleID:
            self.sampleNum.setText(str(sampleID))
    
    def handlePrintPressed(self):
        template = r'C:\Users\simmsk\Desktop\templates\culture_worksheet_template.docx'
        dst = template.split('\\')
        dst[len(dst)-1] = 'temp.docx'
        dst = '\\'.join(dst)
        # shutil.copyfile(template, dst)
        document = MailMerge(template)
        document.merge(
            patientName=f'{self.lastName.text()}, {self.firstName.text()}',
            comments=self.comment.toPlainText()
        )
        document.write(dst)
        word = win32.gencache.EnsureDispatch('Word.Application')
        document = word.Documents.Open(dst)
        txt_path = dst.split('.')[0] + '.html'
        document.SaveAs(txt_path, 10)
        document.Close()
        try:
            word.ActiveDocument()
        except Exception:
            word.Quit()
        os.remove(dst)
        # webbrowser.open(txt_path)
        try:
            self.view.showPrintPreview(txt_path)
        except Exception as e:
            print(e)

class AddClinician(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(AddClinician, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Add_New_Clinician.ui", self)
        # Handle 'Add New Clinician' button clicked
        #self.save.clicked.connect(self.handleSavePressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Add New Clinicians functionality
    #def handleSavePressed(self):
        # Save the form and add Clinician to database

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderForm()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
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

class DUWLNav(QWidget):
    def __init__(self, model, view):
        super(DUWLNav, self).__init__()
        self.view = view
        self.model = model
        self.setFixedHeight(250)
        self.setFixedWidth(600)
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
        self.view.showCultureResultForm()

    # Method for 'CAT' button functionality
    def handleCATPressed(self):
        self.view.showCATResultForm()

    # Method for 'DUWL' button functionality
    def handleDUWLPressed(self):
        self.view.showDUWLResultForm()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()


class CultureResultForm(QMainWindow):
    # Class for the Culture Result Form UI
    def __init__(self, model, view):
        super(CultureResultForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Result Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Result_Form.ui", self)
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

class PrintNav(QtWidgets.QWidget):
    def __init__(self, path):
        super(PrintNav, self).__init__()
        self.setWindowTitle('Document Printer')
        self.editor = QtWidgets.QTextEdit(self)
        self.editor.area.setWidgetResizable(True)
        self.editor.textChanged.connect(self.handleTextChanged)
        self.buttonOpen = QtWidgets.QPushButton('Open', self)
        self.buttonOpen.clicked.connect(self.handleOpen)
        self.buttonPrint = QtWidgets.QPushButton('Print', self)
        self.buttonPrint.clicked.connect(self.handlePrint)
        self.buttonPreview = QtWidgets.QPushButton('Preview', self)
        self.buttonPreview.clicked.connect(self.handlePreview)
        layout = QtWidgets.QGridLayout(self)
        layout.addWidget(self.editor, 0, 0, 1, 3)
        layout.addWidget(self.buttonOpen, 1, 0)
        layout.addWidget(self.buttonPrint, 1, 1)
        layout.addWidget(self.buttonPreview, 1, 2)
        self.handleTextChanged()
        if path:
            file = QtCore.QFile(path)
            if file.open(QtCore.QIODevice.ReadOnly):
                stream = QtCore.QTextStream(file)
                text = stream.readAll()
                info = QtCore.QFileInfo(path)
                if info.completeSuffix() == 'html':
                    self.editor.setHtml(text)
                else:
                    self.editor.setPlainText(text)
                file.close()

    def handleOpen(self):
        path = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Open file', '',
            'HTML files (*.html);;Text files (*.txt);; Word Docs (*.docx)')[0]
        if path:
            file = QtCore.QFile(path)
            if file.open(QtCore.QIODevice.ReadOnly):
                stream = QtCore.QTextStream(file)
                text = stream.readAll()
                info = QtCore.QFileInfo(path)
                if info.completeSuffix() == 'html':
                    self.editor.setHtml(text)
                else:
                    self.editor.setPlainText(text)
                file.close()

    def handlePrint(self):
        dialog = QtPrintSupport.QPrintDialog()
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.editor.document().print_(dialog.printer())

    def handlePreview(self):
        dialog = QtPrintSupport.QPrintPreviewDialog()
        dialog.paintRequested.connect(self.editor.print_)
        dialog.exec_()

    def handleTextChanged(self):
        enable = not self.editor.document().isEmpty()
        self.buttonPrint.setEnabled(enable)
        self.buttonPreview.setEnabled(enable)

class WebView(QWebEngineView):
    def __init__(self, onClose):
        super(WebView, self).__init__()
        self.onClose = onClose

    def closeEvent(self, event):
        self.onClose()
        # event.accept()