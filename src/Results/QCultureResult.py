from PyQt5.uic import loadUi
from pathlib import Path
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem
import json
from mailmerge import MailMerge
from docxtpl import DocxTemplate
from PyQt5.QtCore import QDate, QTimer, QThread
from PyQt5.QtGui import QIcon

from Utility.QAdminLogin import QAdminLogin
from Utility.QIndexedComboBox import QIndexedComboBox
from Utility.QPrefixGraph import QPrefixGraph


class QCultureResult(QMainWindow):
    def __init__(self, model, view):
        super(QCultureResult, self).__init__()
        self.view = view
        self.model = model
        self.swap = QPrefixGraph(self.model)
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

    #@throwsViewableException
    def handleAdvancedSearchPressed(self):
        self.view.showAdvancedSearchScreen(self, "cultureResult")

    #@throwsViewableException
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

    ##@throwsViewableException
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
                item = QIndexedComboBox(i, j, self, True)
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
                item = QIndexedComboBox(i, j, self, False)
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
        return super(QCultureResult, self).eventFilter(source, event)

    ##@throwsViewableException
    def updateTable(self, kind, row, column):
        if kind:
            if row < len(self.aerobicTable):
                if column < len(self.aerobicTable[row]):
                    self.aerobicTable[row][column] = self.aeTWid.cellWidget(row, column).currentText() if self.aeTWid.cellWidget(row, column) else self.aerobicTable[row][column]
        else:
            if row < len(self.anaerobicTable):
                if column < len(self.anaerobicTable[row]):
                    self.anaerobicTable[row][column] = self.anTWid.cellWidget(row, column).currentText() if self.anTWid.cellWidget(row, column) else self.anaerobicTable[row][column]

    ##@throwsViewableException
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

    ##@throwsViewableException
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

    #@throwsViewableException
    def addRowAerobic(self):
        self.aeTWid.setRowCount(self.aeTWid.rowCount()+1)
        self.aerobicTable.append([self.swap.get('Aerobic', 'word')[0]])
        bacteria = QIndexedComboBox(self.aeTWid.rowCount()-1, 0, self, True)
        bacteria.installEventFilter(self)
        bacteria.addItems(self.aerobicList)
        self.aeTWid.setCellWidget(self.aeTWid.rowCount()-1, 0, bacteria)
        for i in range(1, self.aeTWid.columnCount()):
            self.aerobicTable[self.aeTWid.rowCount()-1].append('NA')
            options = QIndexedComboBox(self.aeTWid.rowCount()-1, i, self, True)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.aeTWid.setCellWidget(self.aeTWid.rowCount()-1, i, options)
        self.aeTWid.resizeColumnsToContents()

    #@throwsViewableException
    def addRowAnaerobic(self):
        self.anTWid.setRowCount(self.anTWid.rowCount()+1)
        self.anaerobicTable.append([self.swap.get('Anaerobic', 'word')[0]])
        bacteria = QIndexedComboBox(self.anTWid.rowCount()-1, 0, self, False)
        bacteria.installEventFilter(self)
        bacteria.addItems(self.anaerobicList)
        self.anTWid.setCellWidget(self.anTWid.rowCount()-1, 0, bacteria)
        for i in range(1, self.anTWid.columnCount()):
            self.anaerobicTable[self.anTWid.rowCount()-1].append('NA')
            options = QIndexedComboBox(self.anTWid.rowCount()-1, i, self, False)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.anTWid.setCellWidget(self.anTWid.rowCount()-1, i, options)
        self.anTWid.resizeColumnsToContents()

    #@throwsViewableException
    def delRowAerobic(self):
        if self.aeTWid.rowCount() > 1:
            self.aeTWid.setRowCount(self.aeTWid.rowCount()-1)
            self.aerobicTable.pop()

    #@throwsViewableException
    def delRowAnaerobic(self):
        if self.anTWid.rowCount() > 1:
            self.anTWid.setRowCount(self.anTWid.rowCount()-1)
            self.anaerobicTable.pop()

    #@throwsViewableException
    def addColAerobic(self):
        self.aeTWid.setColumnCount(self.aeTWid.columnCount()+1)
        self.aerobicTable[0].append('Growth')
        header = QIndexedComboBox(0, self.aeTWid.columnCount()-1, self, True)
        header.installEventFilter(self)
        header.addItems(self.headers)
        self.aeTWid.setCellWidget(0, self.aeTWid.columnCount()-1, header)
        for i in range(1, self.aeTWid.rowCount()):
            self.aerobicTable[i].append('NA')
            options = QIndexedComboBox(i, self.aeTWid.columnCount()-1, self, True)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.aeTWid.setCellWidget(i, self.aeTWid.columnCount()-1, options)
        self.aeTWid.resizeColumnsToContents()

    #@throwsViewableException
    def addColAnaerobic(self):
        self.anTWid.setColumnCount(self.anTWid.columnCount()+1)
        self.anaerobicTable[0].append('Growth')
        header = QIndexedComboBox(0, self.anTWid.columnCount()-1, self, False)
        header.installEventFilter(self)
        header.addItems(self.headers)
        header.adjustSize()
        self.anTWid.setCellWidget(0, self.anTWid.columnCount()-1, header)
        for i in range(1, self.anTWid.rowCount()):
            self.anaerobicTable[i].append('NA')
            options = QIndexedComboBox(i, self.anTWid.columnCount()-1, self, False)
            options.installEventFilter(self)
            options.addItems(self.options)
            self.anTWid.setCellWidget(i, self.anTWid.columnCount()-1, options)
        self.anTWid.resizeColumnsToContents()

    #@throwsViewableException
    def delColAerobic(self):
        if self.aeTWid.columnCount() > 1:
            self.aeTWid.setColumnCount(self.aeTWid.columnCount()-1)
            for row in self.aerobicTable:
                row.pop()

    #@throwsViewableException
    def delColAnaerobic(self):
        if self.anTWid.columnCount() > 1:
            self.anTWid.setColumnCount(self.anTWid.columnCount()-1)
            for row in self.anaerobicTable:
                row.pop()

    #@throwsViewableException
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
            data = None
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
                    self.model.getCurrUser(),
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
                    self.view.auditor(self.model.getCurrUser(), "Update", self.saID.text(), 'Culture_Result')
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

    #@throwsViewableException
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
    
    #@throwsViewableException
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

    #@throwsViewableException
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

    #@throwsViewableException
    def handleBackPressed(self):
        self.view.showResultEntryNav()

    #@throwsViewableException
    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

    #@throwsViewableException
    def timerEvent(self):
        self.errorMessage.setText("")
        self.errorMessage2.setText("")