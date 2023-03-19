import datetime
from PyQt5 import QtWidgets, QtPrintSupport
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtCore import QUrl, Qt, QDate
from PyQt5.QtWidgets import QApplication, QAction

import win32com.client as win32
import sys, os

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
from Settings.QManageArchives import QManageArchives
from Settings.QManagePrefixes import QManagePrefixes
from Settings.QHistoricResults import QHistoricResults

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


class QView:

    def __init__(self, model, testMode=False):
        self.model = model
        app = QApplication(sys.argv)
        app.setApplicationDisplayName('COMBDb')
        screen = QAdminLogin(model, self)
        self.widget = QtWidgets.QStackedWidget()
        self.widget.addWidget(screen)
        #self.widget.setGeometry(10,10,1000,800)
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

    def showSetFilePathScreen(self):
        self.setFilePathScreen = QFilePath(self.model, self)
        self.setFilePathScreen.show()

    def showErrorScreen(self, message):
        self.setErrorScreen = QError(self.model, self, message)
        self.setErrorScreen.show()

    def showConfirmationScreen(self):
        self.setConfirmationScreen = QConfirmation(self.model, self)
        self.setConfirmationScreen.show()

    def showArchiveReminderScreen(self):
        self.setArchiveReminderScreen = QArchiveReminder(self.model, self)
        self.setArchiveReminderScreen.show()

    def showAdminLoginScreen(self):
        adminLoginScreen = QAdminLogin(self.model, self)
        self.widget.addWidget(adminLoginScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAdminHomeScreen(self):
        adminHomeScreen = QAdminHome(self.model, self)
        self.widget.addWidget(adminHomeScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showQAReportScreen(self):
        qaReportScreen = QQAReport(self.model, self)
        self.widget.addWidget(qaReportScreen)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showSettingsNav(self):
        self.settingsNav = QSettingsNav(self.model, self)
        self.settingsNav.show()

    def showSettingsManageTechnicianForm(self):
        settingsManageTechnicianForm = QManageTechnicians(self.model, self)
        self.widget.addWidget(settingsManageTechnicianForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showEditTechnician(self, id):
        self.settingsEditTechnician = QEditTechnician(self.model, self, id)
        self.settingsEditTechnician.show()

    def showSettingsManageArchivesForm(self):
        settingsManageArchivesForm = QManageArchives(self.model, self)
        self.widget.addWidget(settingsManageArchivesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showSettingsManagePrefixesForm(self):
        settingsManagePrefixesForm = QManagePrefixes(self.model, self)
        self.widget.addWidget(settingsManagePrefixesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showHistoricResultsForm(self):
        historicResultsForm = QHistoricResults(self.model, self)
        self.widget.addWidget(historicResultsForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showRejectionLogForm(self):
        rejectionLogForm = QRejectionLog(self.model, self)
        self.widget.addWidget(rejectionLogForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAdvancedSearchScreen(self, orderForm, selector):
        self.advancedOrderScreen = QAdvancedSearch(self.model, self, orderForm, selector)
        self.advancedOrderScreen.show()

    def showCultureOrderNav(self):
        self.cultureOrderNav = QOrderNav(self.model, self)
        self.cultureOrderNav.show()

    def showCultureOrderForm(self):
        cultureOrderForm = QCultureOrder(self.model, self)
        self.widget.addWidget(cultureOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAddClinicianScreen(self, dropdown):
        self.addClinician = QClinician(self.model, self, dropdown)
        self.addClinician.show()

    def showDUWLNav(self):
        self.duwlNav = QDUWLNav(self.model, self)
        self.duwlNav.show()

    def showDUWLOrderForm(self):
        duwlOrderForm = QDUWLOrder(self.model, self)
        self.widget.addWidget(duwlOrderForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLReceiveForm(self):
        duwlReceiveForm = QDUWLReceive(self.model, self)
        self.widget.addWidget(duwlReceiveForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showResultEntryNav(self):
        self.resultEntryNav = QResultNav(self.model, self)
        self.resultEntryNav.show()

    def showCultureResultForm(self):
        cultureResultForm = QCultureResult(self.model, self)
        self.widget.addWidget(cultureResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showCATResultForm(self):
        catResultForm = QCATResult(self.model, self)
        self.widget.addWidget(catResultForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showDUWLResultForm(self):
        duwlResultForm = QDUWLResult(self.model, self)
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
            self.web.page().print(self.dialog.printer(), QView.passPrintPrompt)

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
        
    def auditor(self, tech, action, type, form):
        date = str(datetime.datetime.now().month) + '-' + str(datetime.datetime.now().year)
        filename = 'audit logs/audit-' + date + '.txt'
        f = open(filename, 'a+')
        f.write(str(tech)+"."+str(action)+"."+str(type)+"."+str(form)+"."+str(datetime.datetime.now())+"\n")
        f.close()

    def passPrintPrompt(boolean):
            pass