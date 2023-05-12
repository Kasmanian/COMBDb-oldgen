from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets, QtPrintSupport, QtCore
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtWidgets import *
import win32com.client as win32
import sys, os, datetime, json
from mailmerge import MailMerge
from docxtpl import DocxTemplate
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings
from PyQt5.QtCore import QUrl, Qt, QDate, QTimer, QThread
from PyQt5.QtGui import QIcon
import bcrypt, math
import re


def passPrintPrompt(boolean):
        pass


class View:
    def __init__(self, model, testMode=False):
        self.model = model
        app = QApplication(sys.argv)
        app.setApplicationDisplayName('COMBDb')
        screen = AdminLoginScreen(model, self)
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

    def showQAReportScreen(self):
        qaReportScreen = QAReportScreen(self.model, self)
        self.widget.addWidget(qaReportScreen)
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

    # def showSettingsManageArchivesForm(self):
    #     settingsManageArchivesForm = SettingsManageArchivesForm(self.model, self)
    #     self.widget.addWidget(settingsManageArchivesForm)
    #     self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showSettingsManagePrefixesForm(self):
        settingsManagePrefixesForm = SettingsManagePrefixesForm(self.model, self)
        self.widget.addWidget(settingsManagePrefixesForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    # def showHistoricResultsForm(self):
    #     historicResultsForm = HistoricResultsForm(self.model, self)
    #     self.widget.addWidget(historicResultsForm)
    #     self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showRejectionLogForm(self):
        rejectionLogForm = RejectionLogForm(self.model, self)
        self.widget.addWidget(rejectionLogForm)
        self.widget.setCurrentIndex(self.widget.currentIndex()+1)

    def showAdvancedSearchScreen(self, orderForm, selector):
        self.advancedOrderScreen = AdvancedOrderScreen(self.model, self, orderForm, selector)
        self.advancedOrderScreen.show()

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
        self.printer = QPrinter(QPrinter.HighResolution)
        self.dialog = QPrintDialog(self.printer)
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
        
    def auditor(self, tech, action, app, form):
        #date = str(datetime.datetime.now().month) + '-' + str(datetime.datetime.now().year)
        date = str(datetime.datetime.now())
        #filename = 'audit logs/audit-' + date + '.txt'
        #f = open(filename, 'a+')
        #f.write(str(tech)+"."+str(action)+"."+str(type)+"."+str(form)+"."+str(datetime.datetime.now())+"\n")
        #f.close()
        self.model.auditor(tech, action, app, form, date)
        return

def throwsViewableException(func):
    def wrap(self, *args, **kwargs):
        try:
            result = func(self, *args[1:], **kwargs)
            return result
        except TypeError:
            result = func(self, *args, **kwargs)
            return result
        except Exception as e:
            return self.view.showErrorScreen(e)
    return wrap

class PrefixGraph():
    def __init__(self, model):
        self.model = model
        self.__nodes__ = {}
        self.populate('Antibiotics')
        self.populate('Anaerobic')
        self.populate('Aerobic')
        self.populate('Growth')
        self.populate('B-Lac')
        self.populate('Susceptibility')
        
    def translate(self, cat, key, on, to):
        inmap = { 'entry': 0, 'prefix': 1, 'word': 2 }
        graph = [ 1, 2, 0 ]
        on = inmap[on]
        to = inmap[to]
        node = self.__nodes__[cat]
        while (graph[on]!=to):
            if key in node[on]: 
                key = node[on][key]
            else: 
                return None
            on = graph[on]
        return node[on][key] if key in node[on] else None
        
    def populate(self, type):
        typeList = self.model.selectPrefixes(type, 'Entry, Prefix, Word')
        typeEntry, typePrefix, typeWord = {}, {}, {}
        for x in typeList:
            typeEntry.update({x[0]:x[1]})
            typePrefix.update({x[1]:x[2]})
            typeWord.update({x[2]:x[0]})
        self.__nodes__[type] =  [typeEntry, typePrefix, typeWord]
        
    def get(self, cat, field):
        inmap = { 'entry': 0, 'prefix': 1, 'word': 2 }
        return list(self.__nodes__[cat][inmap[field]].keys())

    def exists(self, field, item):
        inmap = { 'entry': 0, 'prefix': 1, 'word': 2 }
        for cat in self.__nodes__:
            if item in self.__nodes__[cat][inmap[field]]:
                return True
        return False
        
# example
#pg = PrefixGraph()
#pg.__nodes__['Growth'][0][0] = 'AB'
#pg.__nodes__['Growth'][1]['AB'] = 'Burn'
#pg.__nodes__['Growth'][2]['Burn'] = 0
#print(pg.translate('Growth', 0, 'entry', 'word'))

class SetFilePathScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetFilePathScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Set_File_Path_Form.ui", self)
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.back.clicked.connect(self.handleBackPressed)
        self.browse.clicked.connect(self.handleBrowsePressed)
        self.save.clicked.connect(self.handleSavePressed)
        with open('local.json', 'r+') as JSON:
            self.currDBText = json.load(JSON)
        self.currDB.setText('Current filepath: ' + self.currDBText['DBQ']) if self.currDBText['DBQ'] != "" else self.currDB.setText("Current filepath: None")

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()
        self.close()

    @throwsViewableException
    def handleBrowsePressed(self):
        fname = QFileDialog.getOpenFileName(self, 'Open File', 'C:', 'MS Access Files (*.accdb)')
        self.filePath.setText(fname[0])

    @throwsViewableException
    def handleSavePressed(self):
        with open('local.json', 'r+') as JSON:
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

class SetErrorScreen(QMainWindow):
    def __init__(self, model, view, message):
        super(SetErrorScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Error_Window.ui", self)
        self.ok.clicked.connect(self.handleOKPressed)
        self.errorMessage.setText(str(message))

    @throwsViewableException
    def handleOKPressed(self):
        self.close()

class SetConfirmationScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetConfirmationScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Confirmation_Window.ui", self)
        self.Cancel.clicked.connect(self.handleCancelPressed)

    @throwsViewableException
    def handleCancelPressed(self):
        self.close()

class SetRejectionReasonScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetRejectionReasonScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Rejection_Log_Reason_Form.ui", self)
        self.cancel.clicked.connect(self.handleCancelPressed)
        self.save.clicked.connect(self.handleSavePressed)
    
    @throwsViewableException
    def handleSavePressed(self):
        self.close()

    @throwsViewableException
    def handleCancelPressed(self):
        self.close()

class SetArchiveReminderScreen(QMainWindow):
    def __init__(self, model, view):
        super(SetArchiveReminderScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Archive_Prompt.ui", self)
        self.no.clicked.connect(self.handleNoPressed)
    
    @throwsViewableException
    def handleNoPressed(self):
        self.close()

class AdminLoginScreen(QMainWindow):
    def __init__(self, model, view):
        super(AdminLoginScreen, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Admin_Login.ui", self)
        self.login.setIcon(QIcon('Icon/loginIcon.png'))
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.login.clicked.connect(self.handleLoginPressed)

    @throwsViewableException
    def handleLoginPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if len(self.user.text())==0 or len(self.pswd.text())==0:
            self.errorMessage.setText("Please input all fields")
        else:
            if self.model.techLogin(self.user.text(), self.pswd.text()):
                global currentTech 
                currentTech = list(self.model.currentTech(self.user.text(), 'Entry'))[0]     
                self.view.auditor(currentTech, 'Login', 'COMBDb', 'System')
                self.view.showAdminHomeScreen()
            else:
                self.errorMessage.setText("Invalid username or password")

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

    def event(self, event):
        if event.type() == QtCore.QEvent.KeyPress:
            if event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
                self.handleLoginPressed()
        return super().event(event)

class AdminHomeScreen(QMainWindow):
    def __init__(self, model, view):
        super(AdminHomeScreen, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Admin_Home_Screen.ui", self)
        self.settings.setIcon(QIcon('Icon/settingsIcon.png'))
        self.logout.setIcon(QIcon('Icon/logoutIcon.png'))
        self.cultureOrder.clicked.connect(self.handleCultureOrderFormsPressed)
        self.resultEntry.clicked.connect(self.handleResultEntryPressed)
        self.qaReport.clicked.connect(self.handleQAReportPressed)
        self.settings.clicked.connect(self.handleSettingsPressed)
        self.logout.clicked.connect(self.handleLogoutPressed)
        PrefixGraph(self.model)
        #self.view.auditor(currentTech, "TEST", 'SAMPLE ID', 'TYPE')

    @throwsViewableException
    def handleCultureOrderFormsPressed(self):
        self.view.showCultureOrderNav()

    @throwsViewableException
    def handleResultEntryPressed(self):
        self.view.showResultEntryNav()

    @throwsViewableException
    def handleQAReportPressed(self):
        self.view.showQAReportScreen()

    @throwsViewableException
    def handleSettingsPressed(self):
        self.view.showSettingsNav()

    @throwsViewableException
    def handleLogoutPressed(self):
        self.view.showAdminLoginScreen()

class QAReportScreen(QMainWindow): 
    def __init__(self, model, view):
        super(QAReportScreen, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_QA_Report_Screen.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.find.clicked.connect(self.handleSearchPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.fromDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.toDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))

    @throwsViewableException
    def handleSearchPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        fromDate = self.view.fSlashDate(self.fromDate.date())
        toDate = self.view.fSlashDate(self.toDate.date())
        fromFormat = datetime.datetime.strptime(fromDate, "%m/%d/%Y").date()
        toFormat = datetime.datetime.strptime(toDate, "%m/%d/%Y").date()
        if fromFormat <= toFormat:
            cultureData = self.model.findSamplesQA('Cultures', '[SampleID], [Type], [Clinician], [Tech], [Received], [Reported]', fromDate, toDate)
            catData = self.model.findSamplesQA('CATs', '[SampleID], [Type], [Clinician], [Tech], [Received], [Reported]', fromDate, toDate)
            data = cultureData + catData
            count = 0
            for tup in data:
                tup = list(tup)
                new = []
                new.append(tup[0])
                new.append(tup[1])
                clinician = self.model.findClinician(tup[2])
                new.append(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3]))
                if tup[3] != 0:
                    techName = list(self.model.findTech(tup[3], 'First, Middle, Last'))
                    new.append(techName[0] + ' ' + techName[1] + ' ' + techName[2])
                else:
                    new.append("")
                new.append(self.view.fSlashDate(tup[4])) if tup[4] != None else new.append("None")
                new.append(self.view.fSlashDate(tup[5])) if tup[5] != None else new.append("None")
                data[count] = new
                count += 1
            data = sorted(data, key=lambda x: x[0])
            self.qaReportTable.setRowCount(len(data))
            for i in range(len(data)):
                self.qaReportTable.setItem(i, 0, QTableWidgetItem(str(data[i][0])))
                self.qaReportTable.setItem(i, 1, QTableWidgetItem(data[i][1]))
                self.qaReportTable.setItem(i, 2, QTableWidgetItem(data[i][2]))
                self.qaReportTable.setItem(i, 3, QTableWidgetItem(data[i][3]))
                self.qaReportTable.setItem(i, 4, QTableWidgetItem(data[i][5]))
                if str(data[i][5]) != 'None' and str(data[i][4]) != 'None':
                    numDays = (datetime.datetime.strptime(data[i][5], '%m/%d/%Y').date() - datetime.datetime.strptime(data[i][4], '%m/%d/%Y').date()).days
                else:
                    numDays = 'Still in Culture'
                self.qaReportTable.setItem(i, 5, QTableWidgetItem(str(numDays)))
            self.qaReportTable.sortItems(0,0)
            self.qaReportTable.resizeColumnsToContents()
            self.view.auditor(currentTech, 'Search', 'COMBDb', 'QAReport')
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("From date must come before to date") 

    @throwsViewableException
    def handlePrintPressed(self):
        self.printList = []
        for i in range(self.qaReportTable.rowCount()):
            self.printList.append([str(self.qaReportTable.item(i, 0).text()), self.qaReportTable.item(i, 1).text(), self.qaReportTable.item(i, 2).text(), self.qaReportTable.item(i, 3).text(), self.qaReportTable.item(i, 4).text(), self.qaReportTable.item(i, 5).text()])
        template = str(Path().resolve())+r'\templates\qa_report_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge(
            tech=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
        )
        document.write(dst)
        context = {
            'headers1' : ['Sample ID', 'Type', 'Clincian', 'Tech', 'Date Reported', 'Days in Culture'],
            'servers1' : self.printList
        }
        document = DocxTemplate(dst)
        document.render(context)
        document.save(dst)
        self.view.convertAndPrint(dst)
        self.view.auditor(currentTech, 'Print', 'COMBDb', 'QAReport')

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class SettingsNav(QMainWindow):
    def __init__(self, model, view):
        super(SettingsNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Admin_Settings_Nav.ui", self)
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.technicianSettings.clicked.connect(self.handleTechnicianSettingsPressed)
        #self.manageArchives.clicked.connect(self.handleManageArchivesPressed)
        self.managePrefixes.clicked.connect(self.handleManagePrefixesPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.changeDatabase.clicked.connect(self.handleChangeDatabasePressed)
        #self.historicResults.clicked.connect(self.handleHistoricResultsPressed)
        self.rejectionLog.clicked.connect(self.handleRejectionLogPressed)

    @throwsViewableException
    def handleChangeDatabasePressed(self):
        self.view.showSetFilePathScreen()
        self.close()

    @throwsViewableException
    def handleTechnicianSettingsPressed(self):
        self.view.showSettingsManageTechnicianForm()
        self.close()

    # @throwsViewableException
    # def handleManageArchivesPressed(self):
    #     self.view.showSettingsManageArchivesForm()
    #     self.close()

    @throwsViewableException
    def handleManagePrefixesPressed(self):
        self.view.showSettingsManagePrefixesForm()
        self.close()

    #@throwsViewableException
    #def handleHistoricResultsPressed(self):
        #self.view.showHistoricResultsForm()
        #self.close()

    @throwsViewableException
    def handleRejectionLogPressed(self):
        self.view.showRejectionLogForm()
        self.close()
    
    @throwsViewableException
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

class SettingsManageTechnicianForm(QMainWindow):
    def __init__(self, model, view):
        super(SettingsManageTechnicianForm, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Settings_Manage_Technicians_Form.ui", self)
        self.addTech.setIcon(QIcon('Icon/addClinicianIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.edit.clicked.connect(self.handleEditPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.techTable.itemSelectionChanged.connect(self.handleTechnicianSelected)
        self.activate.clicked.connect(self.handleActivatePressed)
        self.deactivate.clicked.connect(self.handleDeactivatePressed)
        self.addTech.clicked.connect(self.handleAddTechPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.pswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.confirmPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.activate.setEnabled(False)
        self.deactivate.setEnabled(False)
        self.edit.setEnabled(False)
        self.selectedTechnician = []
        self.updateTable()

    @throwsViewableException
    def updateTable(self):
        techs = self.model.selectTechs('Entry, Username, Active')
        self.techTable.setRowCount(0)
        self.techTable.setRowCount(len(techs)) 
        self.techTable.setColumnCount(3)
        self.techTable.setColumnWidth(0,75)
        self.techTable.setColumnWidth(1,150)
        self.techTable.setColumnWidth(2,75)
        for i in range(0, len(techs)):
            self.techTable.setItem(i,0, QTableWidgetItem(str(techs[i][0])))
            self.techTable.setItem(i,1, QTableWidgetItem(techs[i][1]))
            self.techTable.setItem(i,2, QTableWidgetItem(techs[i][2]))

    @throwsViewableException
    def handleEditPressed(self):
        if len(self.selectedTechnician)>0:
            self.view.showEditTechnician(self.selectedTechnician[1])

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def handleTechnicianSelected(self):
        self.activate.setEnabled(True)
        self.deactivate.setEnabled(True)
        self.edit.setEnabled(True)
        self.selectedTechnician = [
            self.techTable.currentRow(), 
            int(self.techTable.item(self.techTable.currentRow(), 0).text()),
            self.techTable.item(self.techTable.currentRow(), 1).text(),
            self.techTable.item(self.techTable.currentRow(), 2).text(),
        ]
        self.tech.setText(self.techTable.item(self.techTable.currentRow(), 1).text())
    
    @throwsViewableException
    def handleActivatePressed(self): #TODO - KEEP ADDING AUDIT LOG FUNCTIONALITY
        if self.selectedTechnician[3] != 'Yes':
            if self.model.toggleTech(self.selectedTechnician[1], 'Yes'):
                self.selectedTechnician[3] = 'Yes'
                self.techTable.item(self.selectedTechnician[0], 2).setText('Yes')
                self.view.auditor(currentTech, 'Activate', self.selectedTechnician[2], 'Settings_Edit_Technician')

    @throwsViewableException
    def handleDeactivatePressed(self):
        if self.selectedTechnician[3] != 'No':
            if self.model.toggleTech(self.selectedTechnician[1], 'No'):
                self.selectedTechnician[3] = 'No'
                self.techTable.item(self.selectedTechnician[0], 2).setText('No')
                self.view.auditor(currentTech, 'Deactivate', self.selectedTechnician[2], 'Settings_Edit_Technician')

    @throwsViewableException
    def handleAddTechPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        self.techTable.clearSelection()
        user = self.user.text()
        if self.pswd.text()==self.confirmPswd.text() and self.pswd.text() and self.confirmPswd.text():
            if self.fName.text() and self.lName.text() and self.user.text():
                if self.model.findTechUsername(self.user.text()) == None:
                    self.model.addTech(self.fName.text(), self.mName.text(), self.lName.text(), self.user.text(), self.pswd.text())
                    self.updateTable()
                    self.handleClearPressed()
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Successfully added technician: " + user)
                    self.view.auditor(currentTech, 'Add', user, 'Settings_Edit_Technician')
                else:
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("A technician with this username already exists")
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("You must have a first name, last name, and username")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Password and confirm password are required and must match")

    @throwsViewableException
    def handleClearPressed(self):
        self.fName.clear()
        self.mName.clear()
        self.lName.clear()
        self.user.clear()
        self.pswd.clear()
        self.confirmPswd.clear()
        self.techTable.clearSelection()
        self.tech.clear()
        self.activate.setEnabled(False)
        self.deactivate.setEnabled(False)
        self.edit.setEnabled(False)

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class SettingsEditTechnician(QMainWindow):
    def __init__(self, model, view, id):
        super(SettingsEditTechnician, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.id = id
        loadUi("UI Screens/COMBdb_Settings_Edit_Technician.ui", self)
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.tech = self.model.findTech(self.id, '[First], [Middle], [Last], [Username], [Password]')
        self.fName.setText(self.tech[0])
        self.mName.setText(self.tech[1])
        self.lName.setText(self.tech[2])
        self.user.setText(self.tech[3])
        self.oldPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.newPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.confirmNewPswd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.save.clicked.connect(self.handleSavePressed)

    @throwsViewableException
    def handleBackPressed(self):
        self.close()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

    @throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
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
                    self.view.auditor(currentTech, 'Edit', self.user.text(), 'Settings_Edit_Technician')
                    self.close()
                else: 
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText('Old password is incorrect')
            else: 
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("New password and confirm new password don't match")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText('Missing required fields')

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

# class SettingsManageArchivesForm(QMainWindow): #TODO - incorporate archiving.
#     def __init__(self, model, view):
#         super(SettingsManageArchivesForm, self).__init__()
#         self.view = view
#         self.model = model
#         loadUi("UI Screens/COMBdb_Settings_Manage_Archives_Form.ui", self)
#         self.save.setIcon(QIcon('Icon/saveIcon.png'))
#         self.home.setIcon(QIcon('Icon/menuIcon.png'))
#         self.back.setIcon(QIcon('Icon/backIcon.png'))
#         self.back.clicked.connect(self.handleBackPressed)
#         self.home.clicked.connect(self.handleReturnToMainMenuPressed)



            # this can be used when archiving bc it calls a macro that turns the AuditLog table into a txt file ... need to then delete the table and recreate it and re-initilize it
            # strDbName = r"C:\\Users\\Hoboburger\\Desktop\\COMB.accdb"
            # ac = win32.Dispatch("Access.Application")
            # ac.Visible = False
        
            
            # ac.OpenCurrentDatabase(strDbName)
            # ac.DoCmd.RunMacro('ExportAuditLogMacro')
            # ac.DoCmd.CloseDatabase  




#     @throwsViewableException
#     def handleBackPressed(self):
#         self.view.showSettingsNav()

#     @throwsViewableException
#     def handleReturnToMainMenuPressed(self):
#         self.view.showAdminHomeScreen()

class SettingsManagePrefixesForm(QMainWindow):
    def __init__(self, model, view):
        super(SettingsManagePrefixesForm, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.ref = PrefixGraph(self.model)
        loadUi("UI Screens/COMBdb_Settings_Manage_Prefixes_Form.ui", self)
        self.add.setIcon(QIcon('Icon/addIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.add.clicked.connect(self.handleAddPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.aeTWid.itemSelectionChanged.connect(lambda: self.handlePrefixSelected("Aerobic"))
        self.anTWid.itemSelectionChanged.connect(lambda: self.handlePrefixSelected("Anaerobic"))
        self.abTWid.itemSelectionChanged.connect(lambda: self.handlePrefixSelected("Antibiotics"))
        self.prefixesTabWidget.currentChanged.connect(self.clearSelection)
        aeHeader = self.aeTWid.horizontalHeader()
        aeHeader.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        anHeader = self.anTWid.horizontalHeader()
        anHeader.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        abHeader = self.abTWid.horizontalHeader()
        abHeader.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.currentPrefix = ""
        self.selectedPrefix = {}
        self.updateTable("Aerobic")
        self.updateTable("Anaerobic")
        self.updateTable("Antibiotics")
        self.save.setEnabled(False)

    @throwsViewableException
    def clearSelection(self):
        if self.prefixesTabWidget.currentIndex() == 0:
            self.anTWid.clearSelection()
            self.abTWid.clearSelection()
        elif self.prefixesTabWidget.currentIndex() == 1:
            self.aeTWid.clearSelection()
            self.abTWid.clearSelection()
        else:
            self.aeTWid.clearSelection()
            self.anTWid.clearSelection()

    @throwsViewableException
    def updateTable(self, type):
        widget = self.aeTWid if type == "Aerobic" else self.anTWid if type == "Anaerobic" else self.abTWid
        prefix = self.model.selectPrefixes(type, 'Prefix, Word')
        widget.setRowCount(0)
        widget.setRowCount(len(prefix))
        widget.setColumnCount(2)
        widget.setColumnWidth(0, 50)
        widget.setColumnWidth(1, 300)
        for i in range(0, len(prefix)):
            widget.setItem(i,0, QTableWidgetItem(prefix[i][0]))
            widget.setItem(i,1, QTableWidgetItem(prefix[i][1]))
        widget.sortItems(0,0)

    @throwsViewableException
    def handlePrefixSelected(self, type):
        widget = self.aeTWid if type == "Aerobic" else self.anTWid if type == "Anaerobic" else self.abTWid
        prefix = widget.item(widget.currentRow(), 0)
        word = widget.item(widget.currentRow(), 1)
        if prefix and word:
            self.selectedPrefix = {prefix.text() : [type, word.text()]}
            self.pName.setText(list(self.selectedPrefix.keys())[0])
            keyList = self.selectedPrefix.get(list(self.selectedPrefix.keys())[0])
            self.type.setCurrentIndex(self.type.findText(keyList[0]))
            if self.type.currentText() == "Antibiotics":
                self.pName.setEnabled(False)
            else:
                self.pName.setEnabled(True)
            self.word.setText(keyList[1])
            self.currentPrefix = self.model.findPrefix(self.pName.text(), 'Entry, Type, Prefix, Word')
            self.type.setEnabled(False)
            self.add.setEnabled(False)
            self.save.setEnabled(True)

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def handleAddPressed(self): 
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        prefix = self.pName.text()
        word = self.word.text()
        type = self.type.currentText()
        if self.type.currentText() and self.pName.text() and self.word.text():
            if not self.ref.exists('prefix', prefix) and not self.ref.exists('word', word):
                self.model.addPrefixes(self.type.currentText(), self.pName.text(), self.word.text())
                self.ref.populate(type)
                self.updateTable(self.type.currentText())
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Successfully added prefix: " + prefix + ":" + word + " to table: " + type)
                self.view.auditor(currentTech, 'Edit', self.user.text(), 'Settings_Edit_Technician')
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("An entry with that prefix or word already exists")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Type, Prefix and Word are required")

    @throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.pName.text() and self.word.text() and self.type.currentText():
            self.selectedPrefix
            prefixChanged = self.pName.text() not in self.selectedPrefix
            wordChanged = self.word.text() != self.selectedPrefix.get(list(self.selectedPrefix.keys())[0])[1]
            prefixPassed = True
            wordPassed = True
            if prefixChanged:
                prefixPassed = not self.ref.exists('prefix', self.pName.text())
            if wordChanged:
                wordPassed = not self.ref.exists('word', self.word.text())
            if prefixPassed and wordPassed:
                self.model.updatePrefixes(
                    self.currentPrefix[0],
                    self.type.currentText(),
                    self.pName.text(),
                    self.word.text()
                )
                self.ref.populate(self.type.currentText())
                self.updateTable(self.type.currentText())
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Successfully Updated Prefix")
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("An entry with that prefix or word already exists")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Type, Prefix and Word are required")

    @throwsViewableException
    def handleClearPressed(self):
        self.aeTWid.clearSelection()
        self.anTWid.clearSelection()
        self.abTWid.clearSelection()
        self.type.setCurrentIndex(0)
        self.pName.clear()
        self.word.clear()
        self.errorMessage.clear()
        self.add.setEnabled(True)
        self.save.setEnabled(False)
        self.pName.setEnabled(True)
        self.type.setEnabled(True)

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

# class HistoricResultsForm(QMainWindow):
#     def __init__(self, model, view):
#         super(HistoricResultsForm, self).__init__()
#         self.view = view
#         self.model = model
#         self.timer = QTimer(self)
#         loadUi("UI Screens/COMBdb_Settings_Historical_Results_Form.ui", self)
#         #self.home.setIcon(QIcon('Icon/menuIcon.png'))
#         self.back.setIcon(QIcon('Icon/backIcon.png'))
#         self.back.clicked.connect(self.handleBackPressed)
#         #self.home.clicked.connect(self.handleReturnToMainMenuPressed)

#     @throwsViewableException
#     def handleBackPressed(self):
#         self.view.showSettingsNav()

class RejectionLogForm(QMainWindow):
    def __init__(self, model, view):
        super(RejectionLogForm, self).__init__()
        self.view = view
        self.model = model
        #self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Settings_Rejection_Log_Form.ui", self)
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.back.clicked.connect(self.handleBackPressed)
        self.print.clicked.connect(self.handlePrintPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        rejectedCulture = self.model.findRejections('Cultures', '[SampleID], [Type], [Clinician], [Rejection Date], [Rejection Reason]')
        rejectedCAT = self.model.findRejections('CATs', '[SampleID], [Type], [Clinician], [Rejection Date], [Rejection Reason]')
        rejectedDUWL = self.model.findRejections('Waterlines', '[SampleID], [Clinician], [Rejection Date], [Rejection Reason]')
        count = 0
        for entry in rejectedDUWL:
            entry = list(entry)
            entry.insert(1, "Waterline")
            entry = tuple(entry)
            rejectedDUWL[count] = entry
            count += 1
        rejections = rejectedCulture + rejectedCAT + rejectedDUWL
        count = 0
        for tup in rejections:
            tup = list(tup)
            new = []
            new.append(tup[0])
            new.append(tup[1])
            clinician = self.model.findClinician(tup[2])
            new.append(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3]))
            new.append(self.view.fSlashDate(tup[3]))
            new.append(tup[4])
            rejections[count] = new
            count += 1
        rejections = sorted(rejections, key=lambda x: x[0])
        header = self.rejLogTable.horizontalHeader()
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        self.rejLogTable.setRowCount(len(rejections))
        for i in range(0, len(rejections)):
            self.rejLogTable.setItem(i, 0, QTableWidgetItem(str(rejections[i][0])))
            self.rejLogTable.setItem(i, 1, QTableWidgetItem(rejections[i][1]))
            self.rejLogTable.setItem(i, 2, QTableWidgetItem(rejections[i][2]))
            self.rejLogTable.setItem(i, 3, QTableWidgetItem(rejections[i][3]))
            self.rejLogTable.setItem(i, 4, QTableWidgetItem(rejections[i][4]))
        self.rejLogTable.sortItems(0,0)
        self.rejLogTable.resizeColumnsToContents()

    @throwsViewableException
    def handlePrintPressed(self):
        self.printList = []
        for i in range(self.rejLogTable.rowCount()):
            self.printList.append([self.rejLogTable.item(i, 0).text(), self.rejLogTable.item(i, 1).text(), self.rejLogTable.item(i, 2).text(), self.rejLogTable.item(i, 3).text(), self.rejLogTable.item(i, 4).text()])
        template = str(Path().resolve())+r'\templates\rejection_log_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge(
            tech=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
        )
        document.write(dst)
        context = {
            'headers1' : ['Sample ID', 'Type', 'Clincian', 'Rejection Date', 'Rejection Reason'],
            'servers1' : self.printList
        }
        document = DocxTemplate(dst)
        document.render(context)
        document.save(dst)
        self.view.convertAndPrint(dst)
        
    @throwsViewableException
    def handleBackPressed(self):
        self.view.showSettingsNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

class CultureOrderNav(QMainWindow):
    def __init__(self, model, view):
        super(CultureOrderNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Culture_Order_Forms_Nav.ui", self)
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.culture.clicked.connect(self.handleCulturePressed)
        self.duwl.clicked.connect(self.handleDUWLPressed)
        self.back.clicked.connect(self.handleBackPressed)

    @throwsViewableException
    def handleCulturePressed(self):
        self.view.showCultureOrderForm()
        self.close()

    @throwsViewableException
    def handleDUWLPressed(self):
        self.view.showDUWLNav()
        self.close()

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

class AdvancedOrderScreen(QMainWindow):
    def __init__(self, model, view, orderForm, selector):
        super(AdvancedOrderScreen, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.orderForm = orderForm
        self.selector = selector
        loadUi("UI Screens/COMBdb_Advanced_Search_Form2.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.add.setIcon(QIcon('Icon/addIcon.png'))
        self.find.clicked.connect(self.handleSearchPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.clear.clicked.connect(self.handleClearPressed)
        self.add.clicked.connect(self.handleAddPressed)
        self.add.setEnabled(False)
        self.searchTable.itemSelectionChanged.connect(lambda: self.handleOrderSelected())
        
        #set clinician table
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)

        if self.selector == "duwlOrder" or self.selector == "duwlReceive" or self.selector == "duwlResult":
            self.fName.setEnabled(False)
            self.fName.setText("Not searchable")
            self.lName.setEnabled(False)
            self.lName.setText("Not searchable")

    @throwsViewableException
    def handleSearchPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.saID.text() != "":
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID must only contain numbers")
                return
            if len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID must contain 6 digits")
                return         
        
        if self.selector == "cultureOrder":
            #Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = self.view.entries[self.clinDrop.currentText()]['db'] if self.clinDrop.currentText() != "" else 0
            if sampleID == 0 and clin == 0 and self.fName.text() == "" and self.lName.text() == "":
                return
            inputs = {"SampleID" : sampleID if sampleID != 0 else None, "First" : self.fName.text() if self.fName.text() != "" else None, "Last" : self.lName.text() if self.lName.text() != "" else None, "Clinician" : clin if clin != 0 else None}
            #Query data, join and sort
            cultures = self.model.findSamples('Cultures', inputs, '[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]')
            cats = self.model.findSamples('CATs', inputs, '[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]')
            self.results = cultures + cats
            self.results = sorted(self.results, key=lambda x: x[0])
            #Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(i, 0, QTableWidgetItem(str(self.results[i][0])))
                self.searchTable.setItem(i, 1, QTableWidgetItem(str(self.results[i][3]) + " " + str(self.results[i][4])))
                clinician = self.model.findClinician(self.results[i][2])
                self.searchTable.setItem(i, 2, QTableWidgetItem(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3])))
                self.searchTable.setItem(i, 3, QTableWidgetItem(str(self.results[i][5])))

        elif self.selector == "duwlOrder":
            #Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = self.view.entries[self.clinDrop.currentText()]['db'] if self.clinDrop.currentText() != "" else 0
            if sampleID == 0 and clin == 0:
                return
            inputs = {"SampleID" : sampleID if sampleID != 0 else None, "Clinician" : clin if clin != 0 else None}
            #Query data, join and sort
            self.results = self.model.findSamples('Waterlines', inputs, '[Clinician], [Comments], [Notes], [Shipped], [Rejection Date], [Rejection Reason], [SampleID]')
            #Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(i, 0, QTableWidgetItem(str(self.results[i][6])))
                self.searchTable.setItem(i, 1, QTableWidgetItem("N/A"))
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(i, 2, QTableWidgetItem(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3])))
                self.searchTable.setItem(i, 3, QTableWidgetItem("Waterline"))

        elif self.selector == "duwlReceive":
            #Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = self.view.entries[self.clinDrop.currentText()]['db'] if self.clinDrop.currentText() != "" else 0
            if sampleID == 0 and clin == 0:
                return
            inputs = {"SampleID" : sampleID if sampleID != 0 else None, "Clinician" : clin if clin != 0 else None}
            #Query data, join and sort
            self.results = self.model.findSamples('Waterlines', inputs, '[Clinician], [Comments], [Notes], [OperatoryID], [Product], [Procedure], [Collected], [Received], [Rejection Date], [Rejection Reason], [SampleID]')
            #Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(i, 0, QTableWidgetItem(str(self.results[i][10])))
                self.searchTable.setItem(i, 1, QTableWidgetItem("N/A"))
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(i, 2, QTableWidgetItem(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3])))
                self.searchTable.setItem(i, 3, QTableWidgetItem("Waterline"))

        elif self.selector == "cultureResult":
            #Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = self.view.entries[self.clinDrop.currentText()]['db'] if self.clinDrop.currentText() != "" else 0
            if sampleID == 0 and clin == 0 and self.fName.text() == "" and self.lName.text() == "": #this will fail for classes that dont use first and last name
                return
            inputs = {"SampleID" : sampleID if sampleID != 0 else None, "First" : self.fName.text() if self.fName.text() != "" else None, "Last" : self.lName.text() if self.lName.text() != "" else None, "Clinician" : clin if clin != 0 else None}
            #Query data
            self.results = self.model.findSamples('Cultures', inputs, '[ChartID], [Clinician], [First], [Last], [Tech], [Collected], [Received], [Reported], [Type], [Direct Smear], [Aerobic Results], [Anaerobic Results], [Comments], [Notes], [Rejection Date], [Rejection Reason], [SampleID]')
            #Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(i, 0, QTableWidgetItem(str(self.results[i][16])))
                self.searchTable.setItem(i, 1, QTableWidgetItem(str(self.results[i][2]) + " " + str(self.results[i][3])))
                clinician = self.model.findClinician(self.results[i][1])
                self.searchTable.setItem(i, 2, QTableWidgetItem(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3])))
                self.searchTable.setItem(i, 3, QTableWidgetItem(str(self.results[i][8])))

        elif self.selector == "catResult":
            #Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = self.view.entries[self.clinDrop.currentText()]['db'] if self.clinDrop.currentText() != "" else 0
            if sampleID == 0 and clin == 0 and self.fName.text() == "" and self.lName.text() == "": #this will fail for classes that dont use first and last name
                return
            inputs = {"SampleID" : sampleID if sampleID != 0 else None, "First" : self.fName.text() if self.fName.text() != "" else None, "Last" : self.lName.text() if self.lName.text() != "" else None, "Clinician" : clin if clin != 0 else None}
            #Query data
            self.results = self.model.findSamples('CATs', inputs, '[Clinician], [First], [Last], [Tech], [Reported], [Type], [Volume (ml)], [Time (min)], [Initial (pH)], [Flow Rate (ml/min)], [Buffering Capacity (pH)], [Strep Mutans (CFU/ml)], [Lactobacillus (CFU/ml)], [Comments], [Notes], [Collected], [Received], [Rejection Date], [Rejection Reason], [SampleID]')
            #Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(i, 0, QTableWidgetItem(str(self.results[i][19])))
                self.searchTable.setItem(i, 1, QTableWidgetItem(str(self.results[i][1]) + " " + str(self.results[i][2])))
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(i, 2, QTableWidgetItem(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3])))
                self.searchTable.setItem(i, 3, QTableWidgetItem(str(self.results[i][5])))

        elif self.selector == "duwlResult":
            #Set search parameters
            sampleID = int(self.saID.text()) if self.saID.text() != "" else 0
            clin = self.view.entries[self.clinDrop.currentText()]['db'] if self.clinDrop.currentText() != "" else 0
            if sampleID == 0 and clin == 0:
                return
            inputs = {"SampleID" : sampleID if sampleID != 0 else None, "Clinician" : clin if clin != 0 else None}
            #Query data, join and sort
            self.results = self.model.findSamples('Waterlines', inputs, '[Clinician], [Bacterial Count], [CDC/ADA], [Reported], [Comments], [Notes], [Rejection Date], [Rejection Reason], [OperatoryID], [SampleID]')
            #Initialize table
            self.searchTable.setRowCount(len(self.results))
            for i in range(0, len(self.results)):
                self.searchTable.setItem(i, 0, QTableWidgetItem(str(self.results[i][9])))
                self.searchTable.setItem(i, 1, QTableWidgetItem("N/A"))
                clinician = self.model.findClinician(self.results[i][0])
                self.searchTable.setItem(i, 2, QTableWidgetItem(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3])))
                self.searchTable.setItem(i, 3, QTableWidgetItem("Waterline"))

        else:
            print("Could not add order")

    # # function that loads data into search table
    # def initializeTable(self, results):
    #     self.searchTable.setRowCount(len(self.results))
    #     for i in range(0, len(results)):
    #         self.searchTable.setItem(i, 0, QTableWidgetItem(str(results[i][0])))
    #         self.searchTable.setItem(i, 1, QTableWidgetItem(str(results[i][3]) + " " + str(results[i][4])))
    #         clinician = self.model.findClinician(results[i][2])
    #         self.searchTable.setItem(i, 2, QTableWidgetItem(self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3])))
    #         self.searchTable.setItem(i, 3, QTableWidgetItem(str(results[i][5])))

    @throwsViewableException
    def handleOrderSelected(self):
        self.add.setEnabled(True)

    @throwsViewableException
    def handleAddPressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        self.sample = self.results[self.searchTable.currentRow()]
        if self.selector != None:
            self.orderForm.handleSearchPressed(self.sample)
        else:
            print("Could not add order")
        self.close()

    @throwsViewableException
    def handleBackPressed(self):
        self.close()

    @throwsViewableException
    def handleClearPressed(self):
        if self.selector == "duwlOrder" or self.selector == "duwlReceive" or self.selector == "duwlResult":
            self.saID.clear()
            self.clinDrop.setCurrentIndex(0)
            self.searchTable.setRowCount(0)
        else:
            self.saID.clear()
            self.fName.clear()
            self.lName.clear()
            self.clinDrop.setCurrentIndex(0)
            self.searchTable.setRowCount(0)

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")
    

class CultureOrderForm(QMainWindow):
    def __init__(self, model, view):
        super(CultureOrderForm, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Culture_Order_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.addClinician.setIcon(QIcon('Icon/addClinicianIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.print.setIcon(QIcon('Icon/printIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.addClinician.clicked.connect(self.handleAddNewClinicianPressed)
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        #self.print.clicked.connect(self.handlePrintPressed)
        self.print.clicked.connect(self.threader)
        self.clear.clicked.connect(self.handleClearPressed)
        #self.print.setEnabled(False)
        self.colDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))    
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    @throwsViewableException
    def threader(self):
        self.thread = QThread()
        if self.handleSavePressed():
            self.thread.started.connect(self.handlePrintPressed)
            self.thread.start()
            self.thread.exit()

    @throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    @throwsViewableException
    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet("background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px")
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet("background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px")
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    @throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        self.rejectedCheckBox.setEnabled(True)
        self.saID.setEnabled(False)
        self.type.setEnabled(False)
        if data == False:
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID may only contain numbers")
                return
            #self.sample = self.advancedSearch.queryData
            self.sample = self.model.findSample('Cultures', int(self.saID.text()), '[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]')
            if self.sample is None:
                self.sample = self.model.findSample('CATs', int(self.saID.text()), '[SampleID], [ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date], [Rejection Reason]')
                if self.sample is None or len(self.saID.text()) != 6:
                    self.handleClearPressed()
                    self.saID.setText('xxxxxx')
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("Sample ID not found")
        else:
            self.sample = data
            self.saID.setText(str(self.sample[0]))
            data = False
        if self.sample is not None:
            if self.sample[11] != None:
                self.rejectionError.setText("(REJECTED)")
                self.rejectedCheckBox.setChecked(True)
                self.handleRejectedPressed()
            self.chID.setText(self.sample[1])
            clinician = self.model.findClinician(self.sample[2])
            clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
            self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
            self.fName.setText(self.sample[3])
            self.lName.setText(self.sample[4])
            self.type.setCurrentIndex(self.type.findText(self.sample[5]))
            self.colDate.setDate(self.view.dtToQDate(self.sample[6]))
            self.recDate.setDate(self.view.dtToQDate(self.sample[7]))
            self.cText.setText(self.sample[8])
            self.nText.setText(self.sample[9])
            self.rejectedMessage.setText(self.sample[11])
            self.msg = self.sample[11]
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
            self.errorMessage.setText("Found previous order: " + self.saID.text())

    @throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "cultureOrder")

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showCultureOrderNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
    
    @throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.fName.text() and self.lName.text() and self.type.currentText() and self.clinDrop.currentText() != "":
            if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                if self.saID.text() == "":
                    self.saID.setText("0")
                self.sample = self.model.findSample('Cultures', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date]')
                if self.sample is None:
                    self.sample = self.model.findSample('CATs', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Type], [Collected], [Received], [Comments], [Notes], [Rejection Date]')
                    if self.sample is None:
                        table = 'CATs' if self.type.currentText()=='Caries' else 'Cultures'
                        #Create a new db entry - either culture or CAT
                        saID = self.view.model.addPatientOrder(
                            table,
                            self.chID.text(),
                            self.view.entries[self.clinDrop.currentText()]['db'],
                            self.fName.text(),
                            self.lName.text(),
                            self.colDate.date(),
                            self.recDate.date(),
                            self.type.currentText(),
                            currentTech,
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                        )
                        if saID:
                            self.saID.setText(str(saID))
                            #self.save.setEnabled(False)
                            self.print.setEnabled(True)
                            self.saID.setEnabled(False)
                            self.type.setEnabled(False)
                            self.rejectedCheckBox.setEnabled(True)
                            self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                            self.errorMessage.setText("Successfully saved order: " + str(self.saID.text()))
                            self.view.auditor(currentTech, "Create", self.saID.text(), self.type.currentText() + '_Order')
                            return True
                    else: #Update existing CAT Order
                        if not self.saID.isEnabled() and not self.type.isEnabled():
                            if self.rejectedCheckBox.isChecked() and self.sample[9] is None:
                                rejDate = QDate.currentDate()
                            elif self.rejectedCheckBox.isChecked() and self.sample[9] is not None:
                                rejDate = self.view.dtToQDate(self.sample[9])
                            else:
                                rejDate = None
                            self.model.updateCultureOrder(
                                "CATs",
                                int(self.saID.text()),
                                self.chID.text(),
                                self.view.entries[self.clinDrop.currentText()]['db'],
                                self.fName.text(),
                                self.lName.text(),
                                self.colDate.date(),
                                self.recDate.date(),
                                self.type.currentText(),
                                currentTech,
                                self.cText.toPlainText(),
                                self.nText.toPlainText(),
                                rejDate,
                                self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None
                            )
                            #self.view.showConfirmationScreen("Are you sure you want to update an existing culture order?")
                            self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                            self.errorMessage.setText("Existing CAT Order Updated: " + str(self.saID.text())) 
                            self.view.auditor(currentTech, "Update", self.saID.text(), self.type.currentText() + '_Order')
                            return True 
                        else: 
                            self.handleClearPressed()
                            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                            self.errorMessage.setText("Please search order to edit it")
                else: #Update existing Culture Order
                    if not self.saID.isEnabled() and not self.type.isEnabled():
                        if self.rejectedCheckBox.isChecked() and self.sample[9] is None:
                                rejDate = QDate.currentDate()
                        elif self.rejectedCheckBox.isChecked() and self.sample[9] is not None:
                            rejDate = self.view.dtToQDate(self.sample[9])
                        else:
                            rejDate = None
                        self.model.updateCultureOrder(
                            "Cultures",
                            int(self.saID.text()),
                            self.chID.text(),
                            self.view.entries[self.clinDrop.currentText()]['db'],
                            self.fName.text(),
                            self.lName.text(),
                            self.colDate.date(),
                            self.recDate.date(),
                            self.type.currentText(),
                            currentTech,
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                            rejDate,
                            self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None
                        )
                        #self.view.showConfirmationScreen("Are you sure you want to update an existing culture order?")
                        self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear() 
                        self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                        self.errorMessage.setText("Existing Culture Order Updated: " + str(self.saID.text()))
                        self.view.auditor(currentTech, "Update", self.saID.text(), self.type.currentText() + '_Order')
                        return True
                    else: 
                        self.handleClearPressed()
                        self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                        self.errorMessage.setText("Please search order to edit it")
                        return False
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter reason for rejection")
                return False
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("* Denotes Required Fields")
            return False
        
    @throwsViewableException
    def handlePrintPressed(self): 
        if self.type.currentText()!='Caries':
            template = str(Path().resolve())+r'\templates\culture_worksheet_template4.docx'
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
                notes=self.nText.toPlainText(),
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            document.write(dst)
            self.view.convertAndPrint(dst)
        else:
            template = str(Path().resolve())+r'\templates\cat_worksheet_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
            clinician=self.clinDrop.currentText().split(', ')
            document.merge(
                saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:]}',
                received=self.recDate.date().toString(),
                chartID=self.chID.text(),
                clinicianName = clinician[1] + " " + clinician[0],
                patientName=f'{self.lName.text()}, {self.fName.text()}',
                techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
            )
            document.write(dst)
            self.view.convertAndPrint(dst)
        return

    @throwsViewableException
    def handleClearPressed(self):
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
        #self.print.setEnabled(False)
        self.clear.setEnabled(True)
        self.errorMessage.setText("")
        self.tabWidget.setCurrentIndex(0)
        self.saID.setEnabled(True)
        self.type.setEnabled(True)
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.handleRejectedPressed()

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class AddClinician(QMainWindow):
    def __init__(self, model, view, dropdown):
        super(AddClinician, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        self.dropdown = dropdown
        loadUi("UI Screens/COMBdb_Add_New_Clinician.ui", self)
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.clear.clicked.connect(self.handleClearPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clinDrop.currentIndexChanged.connect(self.selectedClinician)
        self.enrollDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))

    def selectedClinician(self):
        if self.clinDrop.currentText() == "":
            return
        else:
            clinician = self.model.findClinicianFull(self.view.entries[self.clinDrop.currentText()]['db'])
            self.title.setCurrentIndex(self.title.findText(clinician[0]))
            self.fName.setText(clinician[1])
            self.lName.setText(clinician[2])
            self.address1.setText(clinician[6])
            self.address2.setText(clinician[7])
            self.city.setText(clinician[8])
            self.state.setCurrentIndex(self.state.findText(clinician[9]))
            self.zip.setText(clinician[10])
            self.phone.setText(clinician[3])
            self.fax.setText(clinician[4])
            self.email.setText(clinician[11])
            self.enrollDate.setDate(self.view.dtToQDate(clinician[12]))
            self.designation.setText(clinician[5])
            self.cText.setText(clinician[13])

    def handleSavePressed(self): #Incorporate validation to make sure clinician is actually added to DB
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        title, first, last = "", "", ""
        if self.fName.text() and self.lName.text() and self.address1.text() and self.city.text() and self.state.currentText() and self.zip.text():
            #self.sample = self.model.findClinicianFull(self.view.entries[self.clinDrop.currentText()]['db'])
            if self.clinDrop.currentText() == "": #and self.sample is None:
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
                title = self.title.currentText()
                first = self.fName.text()
                last = self.lName.text()
                self.view.setClinicianList()
                self.clinDrop.clear()
                self.clinDrop.addItem("")
                self.clinDrop.addItems(self.view.names)
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("New clinician added: " + title + " " + first + " " + last)
                self.view.auditor(currentTech, "Create", title + '' + first + ' ' + last, 'Clinician')
            else:
                self.model.updateClinician(
                    self.view.entries[self.clinDrop.currentText()]['db'],
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
                    self.cText.toPlainText()
                )
                title = self.title.currentText()
                first = self.fName.text()
                last = self.lName.text()
                self.view.setClinicianList()
                self.clinDrop.clear()
                self.clinDrop.addItem("")
                self.clinDrop.addItems(self.view.names)
                self.handleClearPressed()
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Updated Existing Clinician: " + title + " " + first + " " + last)
                self.view.auditor(currentTech, "Update", title + '' + first + ' ' + last, 'Clinician')
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("* Denotes Required Fields")

    @throwsViewableException
    def handleBackPressed(self):
        self.close()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

    @throwsViewableException
    def handleClearPressed(self):
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
        self.clinDrop.setCurrentIndex(0)

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class DUWLNav(QMainWindow):
    def __init__(self, model, view):
        super(DUWLNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_DUWL_Nav.ui", self)
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.orderCulture.clicked.connect(self.handleOrderCulturePressed)
        self.receivingCulture.clicked.connect(self.handleReceivingCulturePressed)
        self.back.clicked.connect(self.handleBackPressed)

    @throwsViewableException
    def handleOrderCulturePressed(self):
        self.close()
        self.view.showDUWLOrderForm()

    @throwsViewableException
    def handleReceivingCulturePressed(self):
        self.close()
        self.view.showDUWLReceiveForm()

    @throwsViewableException
    def handleBackPressed(self):
        self.close()
        self.view.showCultureOrderNav()

class DUWLOrderForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLOrderForm, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_DUWL_Order_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.addClinician.setIcon(QIcon('Icon/addClinicianIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.print.setIcon(QIcon('Icon/printIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clearAll.setIcon(QIcon('Icon/clearAllIcon.png'))
        self.remove.setIcon(QIcon('Icon/removeIcon.png'))
        self.currentKit = 1
        self.kitList = []
        self.printList = {}
        self.kitNum.setText('1')
        self.numOrders.setValue(1)
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.shipDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
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
        self.row.setRange(1, 10)
        self.col.setRange(1, 3)
        self.kitTWid.setHorizontalHeaderLabels(['Sample ID'])
        header = self.kitTWid.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = ""

    @throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "duwlOrder")

    @throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet("background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px")
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet("background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px")
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    @throwsViewableException
    def activateRemove(self):
        self.remove.setEnabled(True)

    @throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if data == False:
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID may only contain numbers")
                return
            if len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID must contain 6 digits")
                return
            self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Clinician], [Comments], [Notes], [Shipped], [Rejection Date], [Rejection Reason]')
            if self.sample is None:
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID not found")
                return
        else:
            self.sample = data
            data = False
            self.saID.setText(str(self.sample[6]))
            
        saID = int(self.saID.text())
        saIDCheck = str(saID)[0:2]+ "-" +str(saID)[2:]
        kitListValues = [value for elem in self.kitList for value in elem.values()]
        if saIDCheck not in kitListValues:
            if self.sample is not None:
                self.saID.setEnabled(False)
                if self.sample[5] != None:
                    self.rejectionError.setText("(REJECTED)")
                    self.rejectedCheckBox.setChecked(True)
                    self.handleRejectedPressed()
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
                self.cText.setText(self.sample[1])
                self.nText.setText(self.sample[2])
                self.shipDate.setDate(self.view.dtToQDate(self.sample[3]))
                self.rejectedMessage.setText(self.sample[5])
                self.msg = self.sample[5]
                self.rejectedCheckBox.setEnabled(True)
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Found previous order: " + str(saID))
        else:
            self.saID.setText('xxxxxx')
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("This DUWL Order has already been added")
            return     

    @throwsViewableException
    def handleAddClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showDUWLNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.clinDrop.currentText():
            if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                if self.saID.text() == "":
                    self.saID.setText("0")
                self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Clinician], [Comments], [Notes], [Shipped], [Rejection Date], [Rejection Reason], [Tech]')
                if self.sample is None:                
                    numOrders = 1 if int(self.numOrders.text()) == None else int(self.numOrders.text())
                    for x in range(numOrders):
                        saID = self.view.model.addWaterlineOrder(
                            self.view.entries[self.clinDrop.currentText()]['db'],
                            self.shipDate.date(),
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                            currentTech
                        )
                        if saID: 
                            self.saID.setText(str(saID))
                            self.kitList.append({
                                'sampleID': f'{str(saID)[0:2]}-{str(saID)[2:]}',
                                'clinician': self.clinDrop.currentText().split(',')[0],
                                'opID': 'Operatory ID: ______________________',
                                'agent': 'Cleaning Agent:  ____________________',
                                'collected': 'Collection Date: _________'
                            })
                            self.printList[str(saID)] = self.currentKit-1
                            self.currentKit = len(self.kitList)+1
                            self.kitNum.setText(str(self.currentKit))
                    self.handleClearPressed()
                    #self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Created New DUWL Order: " + str(saID)) 
                    self.view.auditor(currentTech, "Create", self.saID.text(), 'DUWL_Order')
                else:
                    if not self.saID.isEnabled():
                        sampleID = self.saID.text()
                        if self.rejectedCheckBox.isChecked() and self.sample[4] is None:
                            rejDate = QDate.currentDate()
                        elif self.rejectedCheckBox.isChecked() and self.sample[4] is not None:
                            rejDate = self.view.dtToQDate(self.sample[4])
                        else:
                            rejDate = None
                        saID = self.model.updateWaterlineOrder(
                            int(self.saID.text()),
                            self.view.entries[self.clinDrop.currentText()]['db'],
                            self.shipDate.date(),
                            self.cText.toPlainText(),
                            self.nText.toPlainText(),
                            rejDate,
                            self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None,
                            currentTech
                        )
                        if saID:
                            self.saID.setText(self.saID.text())
                            self.kitList.append({
                                'sampleID': f'{str(self.saID.text())[0:2]}-{str(self.saID.text())[2:]}',
                                'clinician': self.clinDrop.currentText().split(',')[0],
                                'opID': 'Operatory ID: ______________________',
                                'agent': 'Cleaning Agent:  ____________________',
                                'collected': 'Collection Date: _________'
                            })
                            self.printList[self.saID.text()] = self.currentKit-1
                            self.currentKit = len(self.kitList)+1
                            self.kitNum.setText(str(self.currentKit))
                        self.handleClearPressed()
                        #self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                        self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                        self.errorMessage.setText("Existing DUWL Order Updated: " + sampleID)  
                        self.view.auditor(currentTech, "Update", sampleID, 'DUWL_Order')
                    else:
                        self.handleClearPressed()
                        self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                        self.errorMessage.setText("Please search order to edit it")
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter reason for rejection")            
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Please select a clinician")

    @throwsViewableException
    def handleClearPressed(self):
        self.kitNum.setText(str(self.currentKit))
        self.saID.clear()
        self.cText.clear()
        self.nText.clear()
        self.numOrders.setValue(1)
        self.save.setEnabled(True)
        self.clear.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
        self.errorMessage.setText("")
        self.tabWidget.setCurrentIndex(0)
        self.shipDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.saID.setEnabled(True)
        self.handleRejectedPressed()
        self.updateTable()

    @throwsViewableException
    def handleClearAllPressed(self):
        self.kitList.clear()
        self.currentKit = 1
        self.kitNum.setText("1")
        self.printList.clear()
        self.updateTable()
        self.save.setEnabled(True)

    @throwsViewableException
    def handleRemovePressed(self):
        del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
        del self.printList[self.kitTWid.currentItem().text()]
        count = 0
        for key in self.printList.keys():
            self.printList[key] = count
            count += 1
        self.updateTable()
        self.currentKit = len(self.kitList)+1
        self.kitNum.setText(str(self.currentKit))
        self.remove.setEnabled(False)

    def updateTable(self):
        self.kitTWid.setRowCount(len(self.printList.keys()))
        count = 0
        for item in self.printList.keys():
            self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
            count += 1
        if len(self.printList.keys())>0:
            self.print.setEnabled(True)
        else:
            self.print.setEnabled(False)

    @throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve())+r'\templates\duwl_labels.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        x = int(self.row.value()-1)*3+int(self.col.value())-1
        numRows = math.ceil((len(self.kitList)+x)/3)
        labelList = [None]*numRows
        k = 0
        keys = ['sampleID', 'clinician', 'opID', 'agent', 'collected']
        for i in range(0, numRows):
            labelList[i] = {}
            for j in range(0, 3):
                for key in keys:
                    if x>0:
                        labelList[i][key+str(j+1)] = None
                    else:
                        labelList[i][key+str(j+1)] = None if k>= len(self.kitList) else self.kitList[k][key]
                k = k if x>0 else k+1
                x-=1
        document.merge_rows('sampleID1', labelList)
        document.write(dst)
        self.view.convertAndPrint(dst)

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class DUWLReceiveForm(QMainWindow):
    def __init__(self, model, view):
        super(DUWLReceiveForm, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_DUWL_Receive_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.print.setIcon(QIcon('Icon/printIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clearAll.setIcon(QIcon('Icon/clearAllIcon.png'))
        self.remove.setIcon(QIcon('Icon/removeIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
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
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.clearAll.clicked.connect(self.handleClearAllPressed)
        self.remove.clicked.connect(self.handleRemovePressed)
        self.kitTWid.setColumnCount(1)
        self.kitTWid.setHorizontalHeaderLabels(['Sample ID'])
        header = self.kitTWid.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.kitTWid.itemClicked.connect(self.activateRemove)
        self.print.setEnabled(False)
        self.remove.setEnabled(False)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    @throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "duwlReceive")

    @throwsViewableException
    def activateRemove(self):
        self.remove.setEnabled(True)

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showDUWLNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet("background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px")
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet("background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px")
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    @throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if data == False:
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID may only contain numbers")
                return
            if len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID must contain 6 digits")
                return
            self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Clinician], [Comments], [Notes], [OperatoryID], [Product], [Procedure], [Collected], [Received], [Rejection Date], [Rejection Reason]')
            if self.sample is None:
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID not found")
                return
        else:
            self.sample = data
            data = False
            self.saID.setText(str(self.sample[10]))
                
        saID = int(self.saID.text())
        saIDCheck = str(saID)[0:2]+ "-" +str(saID)[2:]
        kitListValues = [value for elem in self.kitList for value in elem.values()]
        if saIDCheck not in kitListValues:
            if self.sample is not None:
                if self.sample[9] != None:
                    self.rejectionError.setText("(REJECTED)")
                    self.rejectedCheckBox.setChecked(True)
                    self.handleRejectedPressed()
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
                self.cText.setText(self.sample[1])
                self.nText.setText(self.sample[2])
                self.operatory.setText(self.sample[3])
                self.product.setText(self.sample[4])
                self.procedure.setText(self.sample[5])
                self.colDate.setDate(self.view.dtToQDate(self.sample[6]))
                self.recDate.setDate(self.view.dtToQDate(self.sample[7]))
                self.rejectedMessage.setText(self.sample[9])
                self.msg = self.sample[9]
                self.save.setEnabled(True)
                self.saID.setEnabled(False)
                self.rejectedCheckBox.setEnabled(True)
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Found DUWL Order: " + str(saID))
        else:
            self.saID.setText('xxxxxx')
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("This DUWL Order has already been added")
            return

    @throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        saID = int(self.saID.text())
        if self.clinDrop.currentText():
            if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Rejection Date]')
                if self.rejectedCheckBox.isChecked() and self.sample[0] is None:
                    rejDate = QDate.currentDate()
                elif self.rejectedCheckBox.isChecked() and self.sample[0] is not None:
                    rejDate = self.sample[0]
                else:
                    rejDate = None
                if self.model.addWaterlineReceiving(
                    saID,
                    self.operatory.text(),
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.colDate.date(),
                    self.recDate.date(),
                    self.product.text(),
                    self.procedure.text(),
                    self.cText.toPlainText(),
                    self.nText.toPlainText(),
                    rejDate,
                    self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None,
                    currentTech
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
                    self.currentKit = len(self.kitList)+1
                    self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                    self.handleClearPressed()
                    self.save.setEnabled(False)
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Saved DUWL Order: " + str(saID))
                    self.view.auditor(currentTech, "Update", str(saID), 'DUWL_Receive')
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter reason for rejection")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Please select a clinician")

    @throwsViewableException
    def handleClearPressed(self):
        self.saID.clear()
        self.saID.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
        self.cText.clear()
        self.nText.clear()
        self.operatory.clear()
        self.procedure.clear()
        self.product.clear()
        self.save.setEnabled(False)
        self.clear.setEnabled(True)
        self.tabWidget.setCurrentIndex(0)
        self.errorMessage.setText("")
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.handleRejectedPressed()
        self.updateTable()

    @throwsViewableException
    def handleClearAllPressed(self):
        self.kitList.clear()
        self.currentKit = 1
        self.printList.clear()
        self.updateTable()

    @throwsViewableException
    def handleRemovePressed(self):
        del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
        del self.printList[self.kitTWid.currentItem().text()]
        count = 0
        for key in self.printList.keys():
            self.printList[key] = count
            count += 1
        self.updateTable()
        self.currentKit = len(self.kitList)+1
        self.remove.setEnabled(False)

    @throwsViewableException
    def updateTable(self):
        self.kitTWid.setRowCount(len(self.printList.keys()))
        count = 0
        for item in self.printList.keys():
            self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
            count += 1
        if len(self.printList.keys())>0:
            self.print.setEnabled(True)
        else:
            self.print.setEnabled(False)

    @throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve())+r'\templates\pending_duwl_cultures_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge_rows('sampleID', self.kitList)
        document.merge(received=self.view.fSlashDate(self.recDate.date()))
        document.write(dst)
        self.view.convertAndPrint(dst)

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class ResultEntryNav(QMainWindow):
    def __init__(self, model, view):
        super(ResultEntryNav, self).__init__()
        self.view = view
        self.model = model
        loadUi("UI Screens/COMBdb_Result_Entry_Forms_Nav.ui", self)
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.culture.clicked.connect(self.handleCulturePressed)
        self.cat.clicked.connect(self.handleCATPressed)
        self.duwl.clicked.connect(self.handleDUWLPressed)
        self.back.clicked.connect(self.handleBackPressed)

    @throwsViewableException
    def handleCulturePressed(self):
        self.close()
        self.view.showCultureResultForm()

    @throwsViewableException
    def handleCATPressed(self):
        self.close()
        self.view.showCATResultForm()

    @throwsViewableException
    def handleDUWLPressed(self):
        self.close()
        self.view.showDUWLResultForm()

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showAdminHomeScreen()
        self.close()

class CultureResultForm(QMainWindow):
    def __init__(self, model, view):
        super(CultureResultForm, self).__init__()
        self.view = view
        self.model = model
        self.swap = PrefixGraph(self.model)
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_Culture_Result_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.repDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.clear.clicked.connect(self.handleClearPressed)
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        #self.printS.clicked.connect(self.handleDirectSmearPressed)
        self.printS.clicked.connect(lambda: self.threader(0))
        #self.printP.clicked.connect(self.handlePreliminaryPressed)
        self.printP.clicked.connect(lambda: self.threader(1))
        #self.printF.clicked.connect(self.handlePerioPressed)
        self.printF.clicked.connect(lambda: self.threader(2))
        self.save.setEnabled(False)
        self.printP.setEnabled(False)
        self.printF.setEnabled(False)
        self.printS.setEnabled(False)
        self.anTWid.setRowCount(0)
        self.anTWid.setColumnCount(0)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 
        try:
            with open('local.json', 'r+') as JSON:
                data = json.load(JSON)
                self.blacList = data['PrefixToB-Lac'].keys()
                self.growthList = data['PrefixToGrowth'].keys()
                self.susceptibilityList = data['PrefixToSusceptibility'].keys()
                self.headers = ['Growth', 'B-lac']
                self.headerIndexes = { 'Growth': 0, 'B-lac': 1 }
                self.options = ['NA'] + list(self.growthList) + list(self.blacList) + list(self.susceptibilityList)
                self.optionIndexes = { 'NA': 0, 'NI': 1, 'L': 2, 'M': 3, 'H': 4, 'P': 5, 'N': 6, 'S': 7, 'I': 8, 'R': 9 }
            aerobic = self.model.selectPrefixes('Aerobic', '[Prefix], [Word]')
            self.aerobicPrefixes = {}
            self.aerobicBacteria = {}
            self.aerobicList = []
            self.aerobicIndex = {}
            for i in range(0, len(aerobic)):
                self.aerobicPrefixes[aerobic[i][0]] = aerobic[i][1]
                self.aerobicBacteria[aerobic[i][1]] = aerobic[i][0]
                self.aerobicList.append(aerobic[i][1])
                self.aerobicIndex[aerobic[i][1]] = i
            anaerobic = self.model.selectPrefixes('Anaerobic', '[Prefix], [Word]')
            self.anaerobicPrefixes = {}
            self.anaerobicBacteria = {}
            self.anaerobicList = []
            self.anaerobicIndex = {}
            for i in range(0, len(anaerobic)):
                self.anaerobicPrefixes[anaerobic[i][0]] = anaerobic[i][1]
                self.anaerobicBacteria[anaerobic[i][1]] = anaerobic[i][0]
                self.anaerobicList.append(anaerobic[i][1])
                self.anaerobicIndex[anaerobic[i][1]] = i
            antibiotics = self.model.selectPrefixes('Antibiotics', '[Prefix], [Word]')
            self.antibioticPrefixes = {}
            self.antibiotics = {}
            self.antibioticsList = []
            self.antibioticsIndex = {}
            for i in range(0, len(antibiotics)):
                self.antibioticPrefixes[antibiotics[i][0]] = antibiotics[i][1]
                self.antibiotics[antibiotics[i][1]] = antibiotics[i][0]
                self.antibioticsList.append(antibiotics[i][1])
                self.antibioticsIndex[antibiotics[i][1]] = i
                self.headers.append(antibiotics[i][0])
                self.headerIndexes[antibiotics[i][0]] = len(self.headers)-1
            self.addRow1.clicked.connect(self.addRowAerobic)
            self.addRow2.clicked.connect(self.addRowAnaerobic)
            self.delRow1.clicked.connect(self.delRowAerobic)
            self.delRow2.clicked.connect(self.delRowAnaerobic)
            self.addCol1.clicked.connect(self.addColAerobic)
            self.addCol2.clicked.connect(self.addColAnaerobic)
            self.delCol1.clicked.connect(self.delColAerobic)
            self.delCol2.clicked.connect(self.delColAnaerobic)
            self.aerobicTable = self.resultToTable(None, "Aerobic")
            self.anaerobicTable = self.resultToTable(None, "Anaerobic")
            self.initTables()
            self.save.clicked.connect(self.handleSavePressed)
        except Exception as e:
            self.view.showErrorScreen(e)

    def threader(self, arg):
        if arg == 0: ui = self.handleDirectSmearPressed
        elif arg == 1: ui = self.handlePreliminaryPressed
        else: ui = self.handlePerioPressed
        self.thread = QThread()
        if self.handleSavePressed():
            self.thread.started.connect(ui)
            self.thread.start()
            self.thread.exit()

    @throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "cultureResult")

    @throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet("background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px")
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet("background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px")
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    @throwsViewableException
    def initTables(self):
        self.aeTWid.setRowCount(0)
        self.aeTWid.setRowCount(len(self.aerobicTable))
        self.anTWid.setRowCount(0)
        self.anTWid.setRowCount(len(self.anaerobicTable))
        self.aeTWid.setColumnCount(0)
        self.aeTWid.setColumnCount(len(self.aerobicTable[0]))
        self.anTWid.setColumnCount(0)
        self.anTWid.setColumnCount(len(self.anaerobicTable[0]))
        self.aeTWid.setColumnWidth(0,290)
        self.anTWid.setColumnWidth(0,290)
        #aerobic
        self.aeTWid.setItem(0,0, QTableWidgetItem('Bacteria'))
        for i in range(0, len(self.aerobicTable)):
            for j in range(0, len(self.aerobicTable[0])):
                item = IndexedComboBox(i, j, self, True)
                item.installEventFilter(self)
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
                item.installEventFilter(self)
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
        self.aeTWid.resizeColumnsToContents()
        self.anTWid.resizeColumnsToContents()

    def eventFilter(self, source, event):
        if (event.type() == QtCore.QEvent.Wheel and isinstance(source, QtWidgets.QComboBox)):
            return True
        return super(CultureResultForm, self).eventFilter(source, event)

    @throwsViewableException
    def updateTable(self, kind, row, column):
        if kind:
            if row < len(self.aerobicTable):
                if column < len(self.aerobicTable[row]):
                    self.aerobicTable[row][column] = self.aeTWid.cellWidget(row, column).currentText() if self.aeTWid.cellWidget(row, column) else self.aerobicTable[row][column]
        else:
            if row < len(self.anaerobicTable):
                if column < len(self.anaerobicTable[row]):
                    self.anaerobicTable[row][column] = self.anTWid.cellWidget(row, column).currentText() if self.anTWid.cellWidget(row, column) else self.anaerobicTable[row][column]

    @throwsViewableException
    def resultToTable(self, result, type):
        if result is not None:
            result = result.replace(' / ', '|')
            result = result.split('/')
            for x in range(0, len(result)):
                tmp = result[x].replace('|', ' / ')
                result[x] = tmp
            table = [[]]
            for i in range(0, len(result)):
                headers = ['Bacteria']
                bacteria = result[i].split(':')
                entryList = self.swap.get(type, 'entry')
                if bacteria[0].isnumeric() and int(bacteria[0]) in entryList:
                    table.append([self.swap.translate(type, int(bacteria[0]), 'entry', 'word')])
                else:
                    table.append([bacteria[0]])
                antibiotics = bacteria[1].split(';')
                for j in range(0, len(antibiotics)):
                    measures = antibiotics[j].split('=')
                    if len(measures) == 1:
                        measures.append("NA")
                    if i<1: 
                        if measures[0] != 'Growth' and measures[0] != 'B-lac' and measures[0].isnumeric():
                            abPrefix = self.swap.translate('Antibiotics', int(measures[0]), 'entry', 'prefix') 
                            measures[0] = abPrefix
                        headers.append(measures[0])
                    table[i+1].append(measures[1]) 
                if i<1: table[0] = headers
            return table
        else:
            return [['Bacteria','Growth', 'PEN', 'AMP', 'CC', 'TET', 'CEP', 'ERY']]

    @throwsViewableException
    def tableToResult(self, table, type):
        if len(table)>1 and len(table[0])>1:
            result = ''
            for i in range(1, len(table)):
                word = table[i][0]  
                table[i][0] = self.swap.translate(type, word, 'word', 'entry')
                if i>1: result += '/'
                result += f'{table[i][0]}:'
                for j in range(1, len(table[i])):
                    if j>1: result += ';'
                    if table[0][j] in self.swap.get('Antibiotics', 'prefix'):
                        tmp = self.swap.translate('Antibiotics', table[0][j], 'prefix', 'entry')
                        table[0][j] = str(tmp)
                    result += f'{table[0][j]}={table[i][j]}'
            return result
        else:
            return None

    @throwsViewableException
    def addRowAerobic(self):
        self.aeTWid.setRowCount(self.aeTWid.rowCount()+1)
        self.aerobicTable.append([self.swap.get('Aerobic', 'word')[0]])
        bacteria = IndexedComboBox(self.aeTWid.rowCount()-1, 0, self, True)
        bacteria.installEventFilter(self)
        bacteria.addItems(self.aerobicList)
        self.aeTWid.setCellWidget(self.aeTWid.rowCount()-1, 0, bacteria)
        for i in range(1, self.aeTWid.columnCount()):
            self.aerobicTable[self.aeTWid.rowCount()-1].append('NA')
            options = IndexedComboBox(self.aeTWid.rowCount()-1, i, self, True)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.aeTWid.setCellWidget(self.aeTWid.rowCount()-1, i, options)
        self.aeTWid.resizeColumnsToContents()

    @throwsViewableException
    def addRowAnaerobic(self):
        self.anTWid.setRowCount(self.anTWid.rowCount()+1)
        self.anaerobicTable.append([self.swap.get('Anaerobic', 'word')[0]])
        bacteria = IndexedComboBox(self.anTWid.rowCount()-1, 0, self, False)
        bacteria.installEventFilter(self)
        bacteria.addItems(self.anaerobicList)
        self.anTWid.setCellWidget(self.anTWid.rowCount()-1, 0, bacteria)
        for i in range(1, self.anTWid.columnCount()):
            self.anaerobicTable[self.anTWid.rowCount()-1].append('NA')
            options = IndexedComboBox(self.anTWid.rowCount()-1, i, self, False)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.anTWid.setCellWidget(self.anTWid.rowCount()-1, i, options)
        self.anTWid.resizeColumnsToContents()

    @throwsViewableException
    def delRowAerobic(self):
        if self.aeTWid.rowCount() > 1:
            self.aeTWid.setRowCount(self.aeTWid.rowCount()-1)
            self.aerobicTable.pop()

    @throwsViewableException
    def delRowAnaerobic(self):
        if self.anTWid.rowCount() > 1:
            self.anTWid.setRowCount(self.anTWid.rowCount()-1)
            self.anaerobicTable.pop()

    @throwsViewableException
    def addColAerobic(self):
        self.aeTWid.setColumnCount(self.aeTWid.columnCount()+1)
        self.aerobicTable[0].append('Growth')
        header = IndexedComboBox(0, self.aeTWid.columnCount()-1, self, True)
        header.installEventFilter(self)
        header.addItems(self.headers)
        self.aeTWid.setCellWidget(0, self.aeTWid.columnCount()-1, header)
        for i in range(1, self.aeTWid.rowCount()):
            self.aerobicTable[i].append('NA')
            options = IndexedComboBox(i, self.aeTWid.columnCount()-1, self, True)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.aeTWid.setCellWidget(i, self.aeTWid.columnCount()-1, options)
        self.aeTWid.resizeColumnsToContents()

    @throwsViewableException
    def addColAnaerobic(self):
        self.anTWid.setColumnCount(self.anTWid.columnCount()+1)
        self.anaerobicTable[0].append('Growth')
        header = IndexedComboBox(0, self.anTWid.columnCount()-1, self, False)
        header.installEventFilter(self)
        header.addItems(self.headers)
        header.adjustSize()
        self.anTWid.setCellWidget(0, self.anTWid.columnCount()-1, header)
        for i in range(1, self.anTWid.rowCount()):
            self.anaerobicTable[i].append('NA')
            options = IndexedComboBox(i, self.anTWid.columnCount()-1, self, False)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.anTWid.setCellWidget(i, self.anTWid.columnCount()-1, options)
        self.anTWid.resizeColumnsToContents()

    @throwsViewableException
    def delColAerobic(self):
        if self.aeTWid.columnCount() > 1:
            self.aeTWid.setColumnCount(self.aeTWid.columnCount()-1)
            for row in self.aerobicTable:
                row.pop()

    @throwsViewableException
    def delColAnaerobic(self):
        if self.anTWid.columnCount() > 1:
            self.anTWid.setColumnCount(self.anTWid.columnCount()-1)
            for row in self.anaerobicTable:
                row.pop()

    @throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if data == False:
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID may only contain numbers")
                self.errorMessage2.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage2.setText("Sample ID may only contain numbers")
                return
            self.sample = self.model.findSample('Cultures', int(self.saID.text()), '[ChartID], [Clinician], [First], [Last], [Tech], [Collected], [Received], [Reported], [Type], [Direct Smear], [Aerobic Results], [Anaerobic Results], [Comments], [Notes], [Rejection Date], [Rejection Reason]')
            if self.sample is None or len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID not found")
                self.errorMessage2.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage2.setText("Sample ID not found")
        else:
            self.sample = data
            self.saID.setText(str(self.sample[16]))
            data = False
        if self.sample is not None:
            if self.sample[15] != None:
                self.rejectionError.setText("(REJECTED)")
                self.rejectedCheckBox.setChecked(True)
                self.handleRejectedPressed()
            self.chID.setText(self.sample[0])
            clinician = self.model.findClinician(self.sample[1])
            clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
            #self.patientName.setText(self.sample[2] + " " + self.sample[3])
            self.fName.setText(self.sample[2])
            self.lName.setText(self.sample[3])
            self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
            self.recDate.setDate(self.view.dtToQDate(self.sample[6]))
            self.repDate.setDate(self.view.dtToQDate(self.sample[7]))
            self.aerobicTable = self.resultToTable(self.sample[10], "Aerobic")
            self.anaerobicTable = self.resultToTable(self.sample[11], "Anaerobic")
            self.cText.setText(self.sample[12])
            self.nText.setText(self.sample[13])
            self.dText.setText(self.sample[9])
            self.rejectedMessage.setText(self.sample[15])
            self.msg = self.sample[15]
            self.initTables()
            self.save.setEnabled(True)
            self.clear.setEnabled(True)
            self.printP.setEnabled(True)
            self.printF.setEnabled(True)
            self.printS.setEnabled(True)
            self.saID.setEnabled(False)
            self.rejectedCheckBox.setEnabled(True)
            #self.printF.setText(self.sample[8])
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
            self.errorMessage.setText("Found Culture Order: " + self.saID.text())
            self.errorMessage2.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
            self.errorMessage2.setText("Found Culture Order: " + self.saID.text())

    #@throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        aerobic = self.tableToResult(self.aerobicTable, "Aerobic")
        anaerobic = self.tableToResult(self.anaerobicTable, "Anaerobic")
        if self.fName.text() and self.lName.text() and self.clinDrop.currentText() != "":
            if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                if self.model.addCultureResult(
                    int(self.saID.text()),
                    self.chID.text(),
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.fName.text(),
                    self.lName.text(),
                    currentTech,
                    self.repDate.date(),
                    self.sample[8],
                    self.dText.toPlainText(),
                    aerobic,
                    anaerobic,
                    self.cText.toPlainText(),
                    self.nText.toPlainText(),
                    QDate.currentDate() if self.rejectedCheckBox.isChecked() else None,
                    self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None
                ):
                    self.handleSearchPressed(False)
                    #self.save.setEnabled(False)
                    self.clear.setEnabled(True)
                    self.printP.setEnabled(True)
                    self.printF.setEnabled(True)
                    self.printS.setEnabled(True)
                    self.saID.setEnabled(False)
                    self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Saved Culture Result Form: " + self.saID.text())
                    self.errorMessage2.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage2.setText("Saved Culture Result Form: " + self.saID.text())
                    self.view.auditor(currentTech, "Update", self.saID.text(), 'Culture_Result')
                    return True
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter reason for rejection")
                self.errorMessage2.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage2.setText("Please enter reason for rejection")
                return False
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("* Denotes Required Fields")
            self.errorMessage2.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage2.setText("* Denotes Required Fields")
            return False

    @throwsViewableException
    def handleDirectSmearPressed(self):
        self.saID.setEnabled(False)
        template = str(Path().resolve())+r'\templates\culture_smear_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        clinician = self.model.findClinician(self.sample[1])
        document.merge(
            saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
            clinicianName=self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[3]),
            collected=self.view.fSlashDate(self.sample[5]),
            received=self.view.fSlashDate(self.recDate.date()),
            chartID=self.chID.text(),
            patientName=f'{self.lName.text()}, {self.fName.text()}',
            cultureType=self.sample[8],
            comments=self.cText.toPlainText(),
            directSmear=self.dText.toPlainText(),
            techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
        )
        document.write(dst)
        self.view.convertAndPrint(dst)
    
    @throwsViewableException
    def handlePreliminaryPressed(self):
        self.saID.setEnabled(False)
        template = str(Path().resolve())+r'\templates\culture_prelim_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        clinician=self.clinDrop.currentText().split(', ')
        document.merge(
            saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
            collected=self.view.fSlashDate(self.sample[5]),
            received=self.view.fSlashDate(self.recDate.date()),
            reported=self.view.fSlashDate(self.repDate.date()),
            chartID=self.chID.text(),
            clinicianName=clinician[1] + " " + clinician[0],
            patientName=f'{self.lName.text()}, {self.fName.text()}',
            comments=self.cText.toPlainText(),
            cultureType=self.sample[8],
            directSmear=self.dText.toPlainText(),
            techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
        )
        document.write(dst)
        context = {
            'headers' : ['Aerotolerant Bacteria']+self.aerobicTable[0][1:],
            'servers': []
        }
        for i in range(1, len(self.aerobicTable)):
            context['servers'].append(self.aerobicTable[i])
        document = DocxTemplate(dst)
        document.render(context)
        document.save(dst)
        self.view.convertAndPrint(dst)

    def handlePerioPressed(self):
        self.saID.setEnabled(False)
        template = str(Path().resolve())+r'\templates\culture_results_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        clinician=self.clinDrop.currentText().split(', ')
        document.merge(
            saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
            collected=self.view.fSlashDate(self.sample[5]),
            received=self.view.fSlashDate(self.recDate.date()),
            reported=self.view.fSlashDate(self.repDate.date()),
            chartID=self.chID.text(),
            clinicianName=clinician[1] + " " + clinician[0],
            patientName=f'{self.lName.text()}, {self.fName.text()}',
            comments=self.cText.toPlainText(),
            cultureType=self.sample[8],
            directSmear=self.dText.toPlainText(),
            techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.'
        )
        document.write(dst)
        context = {
            'headers1' : ['Aerotolerant Bacteria']+self.aerobicTable[0][1:],
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
        self.view.convertAndPrint(dst)

    @throwsViewableException
    def handleClearPressed(self):
        self.saID.clear()
        self.saID.setEnabled(True)
        #self.patientName.clear()
        self.fName.clear()
        self.lName.clear()
        self.clinDrop.setCurrentIndex(0)
        self.chID.clear()
        self.recDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.repDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.cText.clear()
        self.nText.clear()
        self.dText.clear()
        self.printF.setEnabled(False)
        self.printP.setEnabled(False)
        self.printS.setEnabled(False)
        #self.printF.setText("Result")
        self.save.setEnabled(False)
        self.tabWidget.setCurrentIndex(0)
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.errorMessage.clear()
        self.errorMessage2.clear()
        self.msg = ""
        self.handleRejectedPressed()
        self.aerobicTable = self.resultToTable(None, "Aerobic")
        self.anaerobicTable = self.resultToTable(None, "Anaerobic")
        self.initTables()

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")
        self.errorMessage2.setText("")

class CATResultForm(QMainWindow):
    def __init__(self, model, view):
        super(CATResultForm, self).__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_CAT_Result_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.print.setIcon(QIcon('Icon/printIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
        self.clinDrop.addItems(self.view.names)
        self.volume.setText("0.00")
        self.collectionTime.setText("0.00")
        self.flowRate.setText("0.00")
        self.volume.editingFinished.connect(lambda: self.lineEdited(True))
        self.collectionTime.editingFinished.connect(lambda: self.lineEdited(False))
        self.save.setEnabled(False)
        self.print.setEnabled(False)
        self.repDate.setDate(QDate(self.model.date.year, self.model.date.month, self.model.date.day))
        self.back.clicked.connect(self.handleBackPressed)
        self.home.clicked.connect(self.handleReturnToMainMenuPressed)
        self.save.clicked.connect(self.handleSavePressed)
        self.clear.clicked.connect(self.handleClearPressed)
        #self.print.clicked.connect(self.handlePrintPressed)
        self.print.clicked.connect(self.threader)
        self.find.clicked.connect(self.handleSearchPressed)
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    def threader(self):
        self.thread = QThread()
        if self.handleSavePressed():
            self.thread.started.connect(self.handlePrintPressed)
            self.thread.start()
            self.thread.exit()

    @throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "catResult")

    @throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet("background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px")
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet("background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px")
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    @throwsViewableException
    def lineEdited(self, arg):
        lineEdit = self.volume if arg else self.collectionTime
        pattern = re.compile('^[0-9\.]*$')
        if lineEdit.text() != "" and pattern.match(lineEdit.text()):
            if float(self.collectionTime.text()) != 0:
                vol = float(self.volume.text())
                colTime = float(self.collectionTime.text())
                value = str(vol if arg else colTime)
                rate = round(vol / colTime, 2)
                lineEdit.setText(value)
                self.flowRate.setText(str(rate)) 
                self.errorMessage.setText(None)
            else:
                self.flowRate.setText("0.00")
        else:
            lineEdit.setText("0.00")

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if data == False:
            if not self.saID.text().isdigit():
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID must only contain numbers")
                return
            self.sample = self.model.findSample('CATs', int(self.saID.text()), '[Clinician], [First], [Last], [Tech], [Reported], [Type], [Volume (ml)], [Time (min)], [Initial (pH)], [Flow Rate (ml/min)], [Buffering Capacity (pH)], [Strep Mutans (CFU/ml)], [Lactobacillus (CFU/ml)], [Comments], [Notes], [Collected], [Received], [Rejection Date], [Rejection Reason]')
            if self.sample is None or len(self.saID.text()) != 6:
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID not found")
        else:
            self.sample = data
            self.saID.setText(str(self.sample[19]))
            data = False
        if self.sample is not None:
            if self.sample[18] != None:
                self.rejectionError.setText("(REJECTED)")
                self.rejectedCheckBox.setChecked(True)
                self.handleRejectedPressed()
            clinician = self.model.findClinician(self.sample[0])
            clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
            self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
            self.fName.setText(self.sample[1])
            self.lName.setText(self.sample[2])
            self.repDate.setDate(self.view.dtToQDate(self.sample[4]))
            self.volume.setText(str(self.sample[6]) if self.sample[12] is not None else None)
            self.collectionTime.setText(str(self.sample[7]) if self.sample[12] is not None else None)
            self.initialPH.setText(str(self.sample[8]) if self.sample[12] is not None else None)
            self.flowRate.setText(str(self.sample[9]) if self.sample[12] is not None else None)
            self.bufferingCapacityPH.setText(str(self.sample[10]) if self.sample[12] is not None else None)
            self.strepMutansCount.setText(str(self.sample[11]) if self.sample[12] is not None else None)
            self.lactobacillusCount.setText(str(self.sample[12]) if self.sample[12] is not None else None)
            self.cText.setText(self.sample[13])
            self.nText.setText(self.sample[14])
            self.rejectedMessage.setText(self.sample[18])
            self.msg = self.sample[18]
            self.saID.setEnabled(False)
            self.save.setEnabled(True)
            self.print.setEnabled(True)
            self.clear.setEnabled(True)
            self.rejectedCheckBox.setEnabled(True)
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
            self.errorMessage.setText("Found CAT Order: " + self.saID.text())

    @throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if self.fName.text() and self.lName.text() and self.clinDrop.currentText() != "":
            if float(self.collectionTime.text()) != 0:
                if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                    saID = int(self.saID.text())
                    if self.model.addCATResult(
                        saID,
                        self.view.entries[self.clinDrop.currentText()]['db'],
                        self.fName.text(),
                        self.lName.text(),
                        currentTech,
                        self.repDate.date(),
                        "Caries",
                        float(self.volume.text()),
                        float(self.collectionTime.text()),
                        float(self.flowRate.text()),
                        float(self.initialPH.text()),
                        float(self.bufferingCapacityPH.text()),
                        int(self.strepMutansCount.text()),
                        int(self.lactobacillusCount.text()),
                        self.cText.toPlainText(),
                        self.nText.toPlainText(),
                        QDate.currentDate() if self.rejectedCheckBox.isChecked() else None,
                        self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None
                    ):
                        #self.handleSearchPressed()
                        self.save.setEnabled(True)
                        #self.clear.setEnabled(False)
                        self.print.setEnabled(True)
                        #self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                        self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                        self.errorMessage.setText("Saved CAT Result Form: " + str(saID))
                        self.view.auditor(currentTech, "Update", self.saID.text(), 'CAT_Result')
                        return True
                else:
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                    self.errorMessage.setText("Please enter reason for rejection")
                    return False
            else:
                self.flowRate.setText("x.xx")
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Cannot divide by 0")
                return False
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("* Denotes Required Fields")
            return False
        

    @throwsViewableException
    def handleClearPressed(self):
        self.saID.clear()
        self.saID.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
        self.fName.clear()
        self.lName.clear()
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
        self.save.setEnabled(False)
        self.clear.setEnabled(True)
        self.print.setEnabled(False)
        self.errorMessage.setText("")
        self.tabWidget.setCurrentIndex(0)
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.handleRejectedPressed()

    @throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve())+r'\templates\cat_results_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        clinician = self.clinDrop.currentText().split(', ')
        document.merge(
            saID=f'{self.saID.text()[0:2]}-{self.saID.text()[2:6]}',
            patientName=f'{self.fName.text()} {self.lName.text()}',
            clinicianName=clinician[1] + " " + clinician[0],
            collected=self.view.fSlashDate(self.sample[15]),
            received=self.view.fSlashDate(self.sample[16]),
            flowRate=str(self.flowRate.text()),
            bufferingCapacity=str(self.bufferingCapacityPH.text()),
            smCount='{:.2e}'.format(int(self.strepMutansCount.text())),
            lbCount='{:.2e}'.format(int(self.lactobacillusCount.text())),
            reported=self.view.fSlashDate(self.repDate.date()),
            techName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.',
            comments=self.cText.toPlainText()
        )
        document.write(dst)
        self.view.convertAndPrint(dst)

    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class DUWLResultForm(QMainWindow):
    def __init__(self, model, view):
        super().__init__()
        self.view = view
        self.model = model
        self.timer = QTimer(self)
        loadUi("UI Screens/COMBdb_DUWL_Result_Form.ui", self)
        self.find.setIcon(QIcon('Icon/searchIcon.png'))
        self.find2.setIcon(QIcon('Icon/filterIcon.png'))
        self.save.setIcon(QIcon('Icon/saveIcon.png'))
        self.clear.setIcon(QIcon('Icon/clearIcon.png'))
        self.home.setIcon(QIcon('Icon/menuIcon.png'))
        self.print.setIcon(QIcon('Icon/printIcon.png'))
        self.back.setIcon(QIcon('Icon/backIcon.png'))
        self.clearAll.setIcon(QIcon('Icon/clearAllIcon.png'))
        self.remove.setIcon(QIcon('Icon/removeIcon.png'))
        self.clinDrop.clear()
        self.clinDrop.addItem("")
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
        self.find2.clicked.connect(self.handleAdvancedSearchPressed)
        self.clearAll.clicked.connect(self.handleClearAllPressed)
        self.remove.clicked.connect(self.handleRemovePressed)
        self.kitTWid.setColumnCount(1)
        self.kitTWid.setHorizontalHeaderLabels(['Sample ID'])
        header = self.kitTWid.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.bacterialCount.editingFinished.connect(self.lineEdited)
        self.kitTWid.itemClicked.connect(self.activateRemove)
        self.print.setEnabled(False)
        self.remove.setEnabled(False)
        self.cdcADA.setEnabled(False)
        self.rejectedCheckBox.clicked.connect(self.handleRejectedPressed)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.msg = "" 

    @throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "duwlResult")

    @throwsViewableException
    def handleAddNewClinicianPressed(self):
        self.view.showAddClinicianScreen(self.clinDrop)

    @throwsViewableException
    def handleRejectedPressed(self):
        if self.rejectedCheckBox.isChecked():
            self.rejectedMessage.setStyleSheet("background-color: rgb(255, 255, 255); border-style: solid; border-width: 1px")
            self.rejectedMessage.setPlaceholderText("Reason?")
            self.rejectedMessage.setEnabled(True)
            self.rejectedMessage.setText(self.msg)
        else:
            self.rejectedMessage.setStyleSheet("background-color: rgb(123, 175, 212); border-style: solid; border-width: 0px")
            self.rejectedMessage.setPlaceholderText("")
            self.rejectedMessage.setEnabled(False)
            self.rejectedMessage.clear()

    @throwsViewableException
    def lineEdited(self):
        if self.bacterialCount.text().isdigit():
            if int(self.bacterialCount.text()) < 500:
                self.cdcADA.setCurrentIndex(1)
            else:
                self.cdcADA.setCurrentIndex(2)
        else:
            self.bacterialCount.setText("")
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Bacterial Count may only contain positive integers")

    @throwsViewableException
    def activateRemove(self):
        self.remove.setEnabled(True)

    @throwsViewableException
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    @throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    @throwsViewableException
    def handleSearchPressed(self, data):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        if data == False:
            #Comes in here if searching regularly
            if not self.saID.text().isdigit():
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID must only contain numbers")
                return
            if len(self.saID.text()) != 6:
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID must contains 6 digits")
                return
            self.sample = self.model.findSample('Waterlines', int(self.saID.text()), '[Clinician], [Bacterial Count], [CDC/ADA], [Reported], [Comments], [Notes], [Rejection Date], [Rejection Reason], [OperatoryID]')
            if self.sample is None:
                self.handleClearPressed()
                self.saID.setText('xxxxxx')
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Sample ID not found")
                return
        else:
            #Comes in here if searching by advanced lookup form
            self.sample = data
            data = False
            self.saID.setText(str(self.sample[9]))
           
        #Write data into all fields
        saID = int(self.saID.text())
        saIDCheck = str(saID)[0:2]+ "-" + str(saID)[2:]
        kitListValues = [value for elem in self.kitList for value in elem.values()]
        if saIDCheck not in kitListValues: #Check if data already exist in table
            if self.sample is not None:
                self.saID.setEnabled(False)
                if self.sample[7] != None: #Go in here if the order was rejected
                    self.rejectionError.setText("(REJECTED)")
                    self.rejectedCheckBox.setChecked(True)
                    self.handleRejectedPressed()
                clinician = self.model.findClinician(self.sample[0])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinDrop.setCurrentIndex(self.view.entries[clinicianName]['list']+1)
                self.bacterialCount.setText(str(self.sample[1]) if self.sample[1] else None)
                self.cdcADA.setCurrentIndex(self.meets[self.sample[2]] if self.sample[2] else 0)
                self.repDate.setDate(self.view.dtToQDate(self.sample[3]))
                self.cText.setText(self.sample[4])
                self.nText.setText(self.sample[5])
                self.rejectedMessage.setText(self.sample[7])
                self.msg = self.sample[7]
                self.save.setEnabled(True)
                self.rejectedCheckBox.setEnabled(True)
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                self.errorMessage.setText("Found DUWL Order: " + self.saID.text())
        else:
            self.saID.setText('xxxxxx')
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("This DUWL Order has already been added")  
            return              

    @throwsViewableException
    def handleSavePressed(self):
        self.timer.timeout.connect(self.timerEvent)
        self.timer.start(5000)
        saID = int(self.saID.text())
        if self.clinDrop.currentText() != "":
            if (self.rejectedCheckBox.isChecked() and self.rejectedMessage.text() != "") or not self.rejectedCheckBox.isChecked():
                if self.model.addWaterlineResult(
                    saID,
                    self.view.entries[self.clinDrop.currentText()]['db'],
                    self.repDate.date(),
                    int(self.bacterialCount.text()),
                    self.cdcADA.currentText(),
                    self.cText.toPlainText(),
                    self.nText.toPlainText(),
                    QDate.currentDate() if self.rejectedCheckBox.isChecked() else None,
                    self.rejectedMessage.text() if self.rejectedCheckBox.isChecked() else None,
                    currentTech
                ):
                    self.kitList.append({
                        'sampleID': f'{str(saID)[0:2]}-{str(saID)[2:]}',
                        'operatory': self.sample[8],
                        'count': self.bacterialCount.text(),
                        'cdcADA': self.cdcADA.currentText()
                    })
                    self.printList[str(saID)] = self.currentKit-1
                    self.currentKit = len(self.kitList)+1
                    #self.rejectionError.setText("(REJECTED)") if self.rejectedCheckBox.isChecked() else self.rejectionError.clear()
                    self.handleClearPressed()
                    self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: green")
                    self.errorMessage.setText("Saved DUWL Result Form: " + str(saID)) 
                    self.save.setEnabled(False)
                    self.view.auditor(currentTech, "Update", saID, 'DUWL_Result')
            else:
                self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
                self.errorMessage.setText("Please enter reason for rejection")
        else:
            self.errorMessage.setStyleSheet("font: 12pt 'MS Shell Dlg 2'; color: red")
            self.errorMessage.setText("Please select a clinician")

    @throwsViewableException
    def handleClearPressed(self):
        self.saID.clear()
        self.saID.setEnabled(True)
        self.cText.clear()
        self.nText.clear()
        self.bacterialCount.clear()
        self.cdcADA.setCurrentText(None)
        #self.save.setEnabled(True)
        self.clear.setEnabled(True)
        self.clinDrop.setCurrentIndex(0)
        self.tabWidget.setCurrentIndex(0)
        self.errorMessage.setText("")
        self.rejectedCheckBox.setCheckState(False)
        self.rejectedCheckBox.setEnabled(False)
        self.rejectedMessage.setEnabled(False)
        self.rejectionError.clear()
        self.msg = ""
        self.handleRejectedPressed()
        self.updateTable()

    @throwsViewableException
    def handleClearAllPressed(self):
        self.kitList.clear()
        self.currentKit = 1
        self.printList.clear()
        self.updateTable()

    @throwsViewableException
    def handleRemovePressed(self):
        del self.kitList[self.printList[self.kitTWid.currentItem().text()]]
        del self.printList[self.kitTWid.currentItem().text()]
        count = 0
        for key in self.printList.keys():
            self.printList[key] = count
            count += 1
        self.updateTable()
        self.currentKit = len(self.kitList)+1
        self.remove.setEnabled(False)

    @throwsViewableException
    def updateTable(self):
        self.kitTWid.setRowCount(len(self.printList.keys()))
        count = 0
        for item in self.printList.keys():
            self.kitTWid.setItem(count, 0, QTableWidgetItem(item))
            count += 1
        if len(self.printList.keys())>0:
            self.print.setEnabled(True)
        else:
            self.print.setEnabled(False)

    @throwsViewableException
    def handlePrintPressed(self):
        template = str(Path().resolve())+r'\templates\duwl_results_template.docx'
        dst = self.view.tempify(template)
        document = MailMerge(template)
        document.merge_rows('sampleID', self.kitList)
        clinician = self.model.findClinicianFull(self.sample[0])
        document.merge(
            reported=self.view.fSlashDate(self.repDate.date()),
            clinicianName=self.view.fClinicianNameNormal(clinician[0], clinician[1], clinician[2], clinician[5]),
            designation=clinician[5],
            address=clinician[6],
            address2=clinician[7],
            city=clinician[8],
            state=clinician[9],
            zip=str(clinician[10])
        )
        document.write(dst)
        self.view.convertAndPrint(dst)
    
    @throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")

class IndexedComboBox(QComboBox):
    def __init__(self, row, column, form, kind):
        super(IndexedComboBox, self).__init__()
        self.row = row
        self.column = column
        self.form = form
        self.kind = kind
        self.currentIndexChanged.connect(self.handleCurrentIndexChanged)

    @throwsViewableException
    def handleCurrentIndexChanged(self):
        self.form.updateTable(self.kind, self.row, self.column)