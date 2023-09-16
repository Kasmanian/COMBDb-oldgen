import datetime, json, math, os, re, sys, time, bcrypt
from PyQt5 import QtWidgets, QtPrintSupport
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtCore import QUrl, Qt, QDate
from PyQt5.QtWidgets import QApplication, QAction

import win32com.client as win32

from Orders.QOrderNav import QOrderNav
from Orders.QCultureOrder import QCultureOrder
from Orders.QDUWLNav import QDUWLNav
from Orders.QDUWLOrder import QDUWLOrder
from Orders.QDUWLReceive import QDUWLReceive

from Results.QResultNav import QResultNav
from Results.QCultureResult import QCultureResult
from Results.QCATResult import QCATResult
from Results.QDUWLResult import QDUWLResult

from Settings.QSettingsNav import QSettingsNav
from Settings.QManageTechnicians import QManageTechnicians
from Settings.QEditTechnician import QEditTechnician
from Settings.QManagePrefixes import QManagePrefixes

from Utility.QError import QError
from Utility.QConfirmation import QConfirmation
from Utility.QArchiveReminder import QArchiveReminder
from Utility.QAdminLogin import  QAdminLogin
from Utility.QAdminHome import QAdminHome
from Utility.QQAReport import QQAReport
from Utility.QFilePath import QFilePath
from Utility.QRejectionLog import QRejectionLog
from Utility.QAdvancedSearch import QAdvancedSearch
from Utility.QClinician import QClinician
from Utility.QViewableException import QViewableException

class QView:
    def __init__(self, model):
        self.model = model
        app = QApplication(sys.argv)
        app.setApplicationDisplayName('COMBDb')
        screen = QAdminLogin(model, self)
        self.widget = QtWidgets.QStackedWidget()
        self.widget.addWidget(screen)
        self.widget.showMaximized()
        if not self.model.connect():
            self.showSetFilePathScreen()
        else:
            self.setClinicianList()
        try:
            ret = app.exec()
            if self.model:
                if self.model.db:
                    self.model.db.close()
            sys.exit(ret)
        except Exception as e:
            self.showErrorScreen(e)

    @QViewableException.throwsViewableException
    def showSetFilePathScreen(self):
        self.setFilePathScreen = QFilePath(self.model, self)
        self.setFilePathScreen.show()

    @QViewableException.throwsViewableException
    def showErrorScreen(self, message):
        self.setErrorScreen = QError(self.model, self, message)
        self.setErrorScreen.show()

    @QViewableException.throwsViewableException
    def showConfirmationScreen(self):
        self.setConfirmationScreen = QConfirmation(self.model, self)
        self.setConfirmationScreen.show()

    @QViewableException.throwsViewableException
    def showArchiveReminderScreen(self):
        self.setArchiveReminderScreen = QArchiveReminder(self.model, self)
        self.setArchiveReminderScreen.show()

    @QViewableException.throwsViewableException
    def showAdminLoginScreen(self):
        adminLoginScreen = QAdminLogin(self.model, self)
        self.widget.addWidget(adminLoginScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showAdminHomeScreen(self):
        adminHomeScreen = QAdminHome(self.model, self)
        self.widget.addWidget(adminHomeScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showQAReportScreen(self):
        qaReportScreen = QQAReport(self.model, self)
        self.widget.addWidget(qaReportScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showSettingsNav(self):
        self.settingsNav = QSettingsNav(self.model, self)
        self.settingsNav.show()

    def showSettingsManageTechnicianForm(self):
        settingsManageTechnicianForm = QManageTechnicians(self.model, self)
        self.widget.addWidget(settingsManageTechnicianForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showEditTechnician(self, id):
        self.settingsEditTechnician = QEditTechnician(self.model, self, id)
        self.settingsEditTechnician.show()

    @QViewableException.throwsViewableException
    def showSettingsManagePrefixesForm(self):
        settingsManagePrefixesForm = QManagePrefixes(self.model, self)
        self.widget.addWidget(settingsManagePrefixesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showRejectionLogForm(self):
        rejectionLogForm = QRejectionLog(self.model, self)
        self.widget.addWidget(rejectionLogForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showAdvancedSearchScreen(self, orderForm, selector, hi):
        self.advancedOrderScreen = QAdvancedSearch(self.model, self, orderForm, selector, hi)
        self.advancedOrderScreen.show()

    @QViewableException.throwsViewableException
    def showCultureOrderNav(self):
        self.cultureOrderNav = QOrderNav(self.model, self)
        self.cultureOrderNav.show()

    @QViewableException.throwsViewableException
    def showCultureOrderForm(self):
        cultureOrderForm = QCultureOrder(self.model, self)
        self.widget.addWidget(cultureOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showAddClinicianScreen(self):
        self.addClinician = QClinician(self.model, self)
        self.addClinician.show()

    @QViewableException.throwsViewableException
    def showDUWLNav(self):
        self.duwlNav = QDUWLNav(self.model, self)
        self.duwlNav.show()

    @QViewableException.throwsViewableException
    def showDUWLOrderForm(self):
        duwlOrderForm = QDUWLOrder(self.model, self)
        self.widget.addWidget(duwlOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showDUWLReceiveForm(self):
        duwlReceiveForm = QDUWLReceive(self.model, self)
        self.widget.addWidget(duwlReceiveForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showResultEntryNav(self):
        self.resultEntryNav = QResultNav(self.model, self)
        self.resultEntryNav.show()

    @QViewableException.throwsViewableException
    def showCultureResultForm(self):
        cultureResultForm = QCultureResult(self.model, self)
        self.widget.addWidget(cultureResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showCATResultForm(self):
        catResultForm = QCATResult(self.model, self)
        self.widget.addWidget(catResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showDUWLResultForm(self):
        duwlResultForm = QDUWLResult(self.model, self)
        self.widget.addWidget(duwlResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    @QViewableException.throwsViewableException
    def showPrintPreview(self, wordPath, pdfPath):
        self.web = QWebEngineView()
        self.web.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)
        self.web.setWindowTitle("Print Preview")
        self.web.setContextMenuPolicy(Qt.ActionsContextMenu)
        printAction = QAction("Print", self.web)
        printAction.triggered.connect(lambda: self.showPrintPrompt(wordPath))
        self.web.addAction(printAction)
        self.web.load(QUrl.fromLocalFile(pdfPath))
        self.web.showMaximized()

    @QViewableException.throwsViewableException
    def showPrintPrompt(self, wordPath):
        word = win32.Dispatch("Word.Application")
        word.Documents.Open(wordPath)
        word.ActiveDocument.PrintOut()
        time.sleep(5)
        word.ActiveDocument.Close()
        word.Quit()

    @QViewableException.throwsViewableException
    def convertAndPrint(self, wordPath):
        try:
            word = win32.DispatchEx("Word.Application")
            document = word.Documents.Open(wordPath)
            pdfPath = wordPath.split(".")[0] + ".pdf"
            document.SaveAs(pdfPath, 17)
            self.showPrintPreview(wordPath, pdfPath)
        except Exception as e:
            self.showErrorScreen(e)
        finally:
            if document:
                document.Close()
            if word:
                word.Quit()

    @QViewableException.throwsViewableException
    def tempify(self, path):
        tempPath = path.split("\\")
        tempPath[len(tempPath) - 1] = "temp.docx"
        tempPath = "\\".join(tempPath)
        return tempPath

    @QViewableException.throwsViewableException
    def fClinicianName(self, prefix, first, last, designation):
        em = ""
        comma = ", " if first is not None else ""
        prefix = prefix + " " if prefix is not None else prefix
        return (
            f"{last or em}{comma}{prefix or em}{first or em}"
            if prefix is not None or first is not None or last is not None
            else designation or ""
        )

    @QViewableException.throwsViewableException
    def fClinicianNameNormal(self, prefix, first, last, designation):
        em = ""
        prefix = prefix + " " if prefix is not None else prefix
        first = first + " " if first is not None else first
        return (
            f"{prefix or em}{first or em}{last or em}"
            if prefix is not None or first is not None or last is not None
            else designation or ""
        )

    @QViewableException.throwsViewableException
    def fSlashDate(self, date):
        if isinstance(date, datetime.datetime):
            return date.strftime("%m/%d/%Y")
        else:
            return f"{date.month()}/{date.day()}/{date.year()}"

    @QViewableException.throwsViewableException
    def dtToQDate(self, date):
        return (
            QDate(date.year, date.month, date.day)
            if date is not None
            else QDate(self.model.date.year, self.model.date.month, self.model.date.day)
        )

    @QViewableException.throwsViewableException
    def setClinicianList(self):
        try:
            self.clinicians = self.model.selectClinicians(
                "Entry, Prefix, First, Last, Designation, Phone, Fax, Email, [Address 1], [Address 2], City, State, Zip, Enrolled, Inactive, Comments"
            )
            self.entries = {}
            self.names = []
            for clinician in self.clinicians:
                name = self.fClinicianName(
                    clinician[1], clinician[2], clinician[3], clinician[4]
                )
                self.names.append(name)
                self.entries[name] = {"db": clinician[0]}
            self.names.sort()
            for i in range(0, len(self.names)):
                self.entries[self.names[i]]["list"] = i
        except Exception as e:
            self.showErrorScreen(e)
        
    @QViewableException.throwsViewableException
    def auditor(self, tech, action, app, form):
        date = str(datetime.datetime.now())
        self.model.auditor(tech, action, app, form, date)
        return
