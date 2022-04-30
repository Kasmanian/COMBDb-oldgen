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

    def showSettingsNav(self):
        self.settingsNav = SettingsNav(self.model, self)
        self.settingsNav.show()

    def showSettingsManageTechnicianForm(self):
        settingsManageTechnicianForm = SettingsManageTechnicianForm(self.model, self)
        self.widget.addWidget(settingsManageTechnicianForm)
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
        self.addClinician = AddClinician(self.model, self)
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
        # self.guestLogin.clicked.connect(self.handleGuestLoginPressed)

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
        self.view.showSettingsNav()

    # Method for 'Logout' button functionality
    def handleLogoutPressed(self):
        print('Logged out...')
        self.view.showAdminLoginScreen()


class SettingsNav(QMainWindow):
    # Class for the Settings Nav UI
    def __init__(self, model, view):
        super(SettingsNav, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Admin_Settings_Nav.ui", self)
        # Handle 'Technician Settings' button clicked
        self.technicianSettings.clicked.connect(self.handleTechnicianSettingsPressed)
        # Handle 'Manage Archives' button clicked
        #self.manageArchives.clicked.connect(self.handleManageArchivesPressed)
        # Handle 'Manage Prefixes' button clicked
        #self.managePrefixes.clicked.connect(self.handleManagePrefixesPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)

    # Method for 'Technician Settings' button functionality
    def handleTechnicianSettingsPressed(self):
        self.view.showSettingsManageTechnicianForm()
        self.close()

    # Method for 'Manage Archives' button functionality
    #def handleManageArchivesPressed(self):
        #self.view.showAdminHomeScreen()

    # Method for 'Manage Prefixes' button functionality
    #def handleManagePrefixesPressed(self):
        #self.view.showAdminHomeScreen()

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.close()

class SettingsManageTechnicianForm(QMainWindow):
     # Class for the Settings Screen UI
    def __init__(self, model, view):
        super(SettingsManageTechnicianForm, self).__init__()
        self.view = view
        self.model = model
        # Load the .ui file of the Admin Main Screen 
        loadUi("COMBDb/UI Screens/COMBdb_Settings_Manage_Technicians_Form.ui", self)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

    # Method for 'Back' button functionality
    def handleBackPressed(self):
        self.view.showSettingsNav()

    # Method for 'Return to Main Menu' button functionality
    def handleReturnToMainMenuPressed(self):
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
        # Load the .ui file of the Culture Order Form Screen
        loadUi("COMBDb/UI Screens/COMBdb_Culture_Order_Form.ui", self)
        # Handle 'Add New Clinician' button clicked
        self.addClinician.clicked.connect(self.handleAddNewClinicianPressed)
        # Handle 'Back' button clicked
        self.back.clicked.connect(self.handleBackPressed)
        # Handle 'Menu' button clicked
        self.menu.clicked.connect(self.handleReturnToMainMenuPressed)

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