from tkinter import CENTER, Button
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
import sys

class View(QApplication):
    def __init__(self, model):
        self.model = model
        app = QApplication(sys.argv)
        # Launch application straight to Admin Login Screen
        screen = LoginScreen(model, self)
        self.widget = QtWidgets.QStackedWidget()
        self.widget.addWidget(screen)
        self.widget.setFixedHeight(1200)
        self.widget.setFixedWidth(1600)
        self.widget.show()
        try:
            sys.exit(app.exec())
        except:
            print("Exiting")

    def showGuestLoginScreen(self):
        guestLoginScreen = GuestLoginScreen(self.model, self)
        self.widget.addWidget(guestLoginScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAdminHomeScreen(self):
        adminHomeScreen = AdminHomeScreen(self.model, self)
        self.widget.addWidget(adminHomeScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showGuestHomeScreen(self):
        guestHomeScreen = GuestHomeScreen(self.model, self)
        self.widget.addWidget(guestHomeScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCultureOrderNav(self): # This is the Nav Menu for Culture Orders and below are subsequent screens
        cultureOrderNav = CultureOrderNav(self.model, self)
        self.widget.addWidget(cultureOrderNav)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCultureOrderForm(self):
        cultureOrderForm = CultureOrderForm(self.model, self)
        self.widget.addWidget(cultureOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCATOrderForm(self):
        catOrderForm = CATOrderForm(self.model, self)
        self.widget.addWidget(catOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLOrderForm(self):
        duwlOrderForm = DUWLOrderForm(self.model, self)
        self.widget.addWidget(duwlOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showResultEntryNav(self): # This is the Nav Menu for Result Entry and below are subsequent screens
        resultEntryNav = ResultEntryNav(self.model, self)
        self.widget.addWidget(resultEntryNav)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCultureResultForm(self):
        cultureResultForm = CultureResultForm(self.model, self)
        self.widget.addWidget(cultureResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    #def showPreviousScreen(self):
        #elf.widget.setCurrentIndex(self.widget.currentIndex()-1)


class LoginScreen(QMainWindow):
    # Class for the Login Screen UI
    def __init__(self, model, view):
        super(LoginScreen, self).__init__()
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
        if self.model.adminLogin(self.usrnm.text(), self.pswd.text()):
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

    # Method for 'Culture Order Forms' button functionality
    def handleCultureOrderFormsPressed(self):
        self.view.showCultureOrderNav()

    def handleResultEntryPressed(self):
        self.view.showResultEntryNav()


class GuestHomeScreen(QMainWindow):
    # Class for the Guest Home Screen UI
    def __init__(self, model, view):
        super(GuestHomeScreen, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Guest_Home_Screen.ui", self)


class CultureOrderNav(QMainWindow):
    # Class for the Culture Order Forms Navigation Screen UI
    def __init__(self, model, view):
        super(CultureOrderNav, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Navigation Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Forms_Nav.ui", self)
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
        self.view.showCultureOrderForm()

    # Method for 'CAT' button functionality
    def handleCATPressed(self):
        self.view.showCATOrderForm()

    # Method for 'DUWL' button functionality
    def handleDUWLPressed(self):
        self.view.showDUWLOrderForm()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()
        #self.view.showPreviousScreen()


class CultureOrderForm(QMainWindow):
    # Class for the Culture Order Form UI
    def __init__(self, model, view):
        super(CultureOrderForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)
        # Handle 'Save' button clicked
        self.save.clicked.connect(self.handleSavePressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        #self.view.showPreviousScreen()
    
    # Method for entering new Patient Sample Data
    def handleSavePressed(self):
        print(self.clinicianDropDown.currentText())
        print(self.cultureTypeDropDown.currentText())
        print(self.chartNum.text())
        print(str(self.collectionDate.date()))
        print(str(self.receivedDate.date()))
        print(self.clinicLocationDropDown.currentText())
        print(self.comment.toPlainText())
        self.view.model.addPatientSample(
            self.firstName.text(),
            self.lastName.text(),
            self.clinicianDropDown.currentText(),
            self.cultureTypeDropDown.currentText(),
            self.chartNum.text(),
            str(self.collectionDate.date()),
            str(self.receivedDate.date()),
            self.clinicLocationDropDown.currentText(),
            self.comment.toPlainText()
        )

class CATOrderForm(QMainWindow):
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
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Culture' button functionality
    def handleCulturePressed(self):
        self.view.showCultureResultForm()

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