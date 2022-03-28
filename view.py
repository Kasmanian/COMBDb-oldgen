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

    def showCultureOrderNav(self):
        cultureOrderNav = CultureOrderNav(self.model, self)
        self.widget.addWidget(cultureOrderNav)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCultureOrderForm(self):
        cultureOrderForm = CultureOrderForm(self.model, self)
        self.widget.addWidget(cultureOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showGuestHomeScreen(self):
        guestHomeScreen = GuestHomeScreen(self.model, self)
        self.widget.addWidget(guestHomeScreen)
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
        loadUi("COMBDb/UI Screens/COMBdb_Login.ui", self)
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        # Handle 'Login' button clicked
        self.login.clicked.connect(self.handleLoginPressed)
        self.guestLogin.clicked.connect(self.handleGuestLoginPressed)

    def handleLoginPressed(self):
        # If credential check is successful, display Admin Home Screen
        if self.model.adminLogin(self.usrnm.text(), self.pswd.text().encode('utf-8')):
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
        loadUi("COMBDb/UI Screens/COMBdb_Main_Screen.ui", self)
        # Handle 'Culture Order Forms' button clicked
        self.cultureOrder.clicked.connect(self.handleCultureOrderFormsPressed)

    # Method for 'Culture Order Forms' button functionality
    def handleCultureOrderFormsPressed(self):
        self.view.showCultureOrderNav()


class GuestHomeScreen(QMainWindow):
    # Class for the Guest Home Screen UI
    def __init__(self, model, view):
        super(GuestHomeScreen, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Guest_Main_Screen.ui", self)


class CultureOrderNav(QMainWindow):
    # Class for the Culture Order Forms Navigation Screen UI
    def __init__(self, model, view):
        super(CultureOrderNav, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Culture Order Navigation Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Nav.ui", self)
        # Handle 'Culture' button clicked
        self.culture.clicked.connect(self.handleCulturePressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Culture' button functionality
    def handleCulturePressed(self):
        self.view.showCultureOrderForm()

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

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        #self.view.showPreviousScreen()
