from PyQt5.QtWidgets import QMainWindow, QTableWidgetItem, QComboBox, QDate
from PyQt5.uic import loadUi
from pathlib import Path
from mailmerge import MailMerge
from docxtpl import DocxTemplate
from util import formatSaID, formatChID, formatDate

class QResultCulture(QMainWindow):
    def __init__(self, app):
        super(QResultCulture, self).__init__()
        loadUi('COMBDb/UI Screens/COMBdb_Culture_Result_Form.ui', self)
        self.clinDrop.clear()
        self.clinDrop.addItems(self.app.names)
        self.recDate.setDate(QDate(self.app.date.year, self.app.date.month, self.app.date.day))
        self.repDate.setDate(QDate(self.app.date.year, self.app.date.month, self.app.date.day))
        self.printD.clicked.connect(self.handlePrintD)
        self.printP.clicked.connect(self.handlePrintP)
        self.printF.clicked.connect(self.handlePrintF)
        self.back.clicked.connect(self.handleBack)
        self.home.clicked.connect(self.handleHome)
        self.save.clicked.connect(self.handleSave)
        self.find.clicked.connect(self.handleFind)
        self.addRow1.clicked.connect(lambda: self.addRow(True))
        self.addRow2.clicked.connect(lambda: self.addRow(False))
        self.delRow1.clicked.connect(lambda: self.delRow(True))
        self.delRow2.clicked.connect(lambda: self.delRow(False))
        self.addCol1.clicked.connect(lambda: self.addCol(True))
        self.addCol2.clicked.connect(lambda: self.addCol(False))
        self.delCol1.clicked.connect(lambda: self.delCol(True))
        self.delCol2.clicked.connect(lambda: self.delCol(False))
        self.printD.setEnabled(False)
        self.printP.setEnabled(False)
        self.printS.setEnabled(False)
        self.save.setEnabled(False)

        self.app = app

        try:
            fields = ('[Entry]', '[Prefix]', '[Expansion]')
            tables = ['Aerobics', 'Anaerobics', 'Antibiotics']
            self.items = []
            for table in tables:   
                self.items[table] = self.app.db.select(table, fields, None, 0)
            self.aeTable = self.resultToTable(None)
            self.anTable = self.resultToTable(None)
            self.initTables()
        except Exception as e:
            self.app.showErrorScreen(e)

    def initTables(self):
        try:
            tables = [self.aeTable, self.anTable]
            tbWids = [self.tableWidget, self.tableWidget_2]
            for t in range(0, 2):
                table = tables[t]
                tbWid = tbWids[t]

                tbWid.setRowCount(0)
                tbWid.setRowCount(len(table))
                tbWid.setColumnCount(0)
                tbWid.setColumnCount(len(table))
                tbWid.setColumnWidth(0,300)
                tbWid.setItem(0,0, QTableWidgetItem('Bacteria'))

                for i in range(0, len(table)):
                    for j in range(0, len(table[0])):
                        item = IndexedComboBox(i, j, self, True)
                        if i>0 and j>0:
                            item.addItems(self.options)
                            item.setCurrentIndex(self.optionIndexes[table[i][j]])
                        elif i<1 and j>0:
                            item.addItems(self.headers)
                            item.setCurrentIndex(self.headerIndexes[table[i][j]])
                        elif i>0 and j<1:
                            item.addItems(self.aerobicList)
                            item.setCurrentIndex(self.aerobicIndex[table[i][j]])
                        else: continue
                        tbWid.setCellWidget(i, j, item)
        except Exception as e:
            self.app.showErrorScreen(e)

    def updateTable(self, aerobic, row, column):
        try:
            tables = [self.aeTable, self.tableWidget] if aerobic else [self.anTable, self.tableWidget_2]
            self.tables[0][row][column] = self.tables[1].cellWidget(row, column).currentText() if self.tables[1].cellWidget(row, column) else tables[0][row][column]
        except Exception as e:
            self.app.showErrorScreen(e)

    def resultToTable(self, result):
        if result is not None:
            result = result.split('/')
            table = [[]]
            for i in range(0, len(result)):
                headers = ['Bacteria']
                bacteria = result[i].split(':')
                table.append([bacteria[0]])
                antibiotics = bacteria[1].split(';')
                for j in range(0, len(antibiotics)):
                    measures = antibiotics[j].split('=')
                    if i<1: headers.append(measures[0])
                    table[i+1].append(measures[1])
                if i<1: table[0] = headers
            return table
        else:
            return [['Bacteria','Growth', 'B-lac', 'PEN', 'AMP', 'CC', 'TET', 'CEP', 'ERY']]

    def tableToResult(self, table):
        if len(table)>1 and len(table[0])>1:
            result = ''
            for i in range(1, len(table)):
                if i>1: result += '/'
                result += f'{table[i][0]}:'
                for j in range(1, len(table[i])):
                    if j>1: result += ';'
                    result += f'{table[0][j]}={table[i][j]}'
            return result
        else:
            return None

    def addRow(self, aerobic: bool):
        try:
            bx = 'Alpha-Hemolytic Streptococcus' if aerobic else 'Actinobacillus Actinomycetemcomitians'
            tb = [self.aeTbArr, self.aeTbWid] if aerobic else [self.anTbArr, self.anTbWid]
            tb[1].setRowCount(tb[1].rowCount()+1)
            tb[0].append([bx])
            bacteria = IndexedComboBox(tb[1].rowCount()-1, 0, self, True)
            bacteria.addItems(self.aerobicList)
            tb[1].setCellWidget(tb[1].rowCount()-1, 0, bacteria)
            for i in range(1, tb[1].columnCount()):
                tb[0][tb[1].rowCount()-1].append('NI')
                options = IndexedComboBox(tb[1].rowCount()-1, i, self, True)
                options.addItems(self.options)
                tb[1].setCellWidget(tb[1].rowCount()-1, i, options)
        except Exception as e:
            self.app.showErrorScreen(e)

    def delRow(self, aerobic: bool):
        try:
            tb = [self.aeTbArr, self.aeTbWid] if aerobic else [self.anTbArr, self.anTbWid]
            if tb[1].rowCount() > 1:
                self.tb[1].setRowCount(self.tb[1].rowCount()-1)
                self.tb[0].pop()
        except Exception as e:
            self.app.showErrorScreen(e)

    def addCol(self, aerobic: bool):
        try:
            tb = [self.aeTbArr, self.aeTbWid] if aerobic else [self.anTbArr, self.anTbWid]
            tb[1].setColumnCount(tb[1].columnCount()+1)
            tb[0][0].append('Growth')
            header = IndexedComboBox(0, tb[1].columnCount()-1, self, True)
            header.addItems(self.headers)
            tb[1].setCellWidget(0, tb[1].columnCount()-1, header)
            for i in range(1, tb[1].rowCount()):
                tb[0][i].append('NI')
                options = IndexedComboBox(i, tb[1].columnCount()-1, self, True)
                options.addItems(self.options)
                tb[1].setCellWidget(i, tb[1].columnCount()-1, options)
        except Exception as e:
            self.app.showErrorScreen(e)

    def delCol(self, aerobic: bool):
        try:
            tb = [self.aeTbArr, self.aeTbWid] if aerobic else [self.anTbArr, self.anTbWid]
            if tb[1].columnCount() > 1:
                tb[1].setColumnCount(tb[1].columnCount()-1)
                for row in tb[0]:
                    row.pop()
        except Exception as e:
            self.app.showErrorScreen(e)

    def handleFind(self):
        try:
            if not self.sampleID.text().isdigit():
                self.sampleID.setText('xxxxxx')
                return
            self.sample = self.model.findSample('Cultures', int(self.sampleID.text()), '[ChartID], [Clinician], [First], [Last], [Collected], [Received], [Reported], [Aerobic Results], [Anaerobic Results], [Comments]')
            if self.sample is None:
                self.sampleID.setText('xxxxxx')
            else:
                self.chartNumber.setText(self.sample[0])
                clinician = self.model.findClinician(self.sample[1])
                clinicianName = self.view.fClinicianName(clinician[0], clinician[1], clinician[2], clinician[3])
                self.clinician.setCurrentIndex(self.view.entries[clinicianName]['list'])
                self.receivedDate.setDate(self.view.dtToQDate(self.sample[5]))
                self.dateReported.setDate(self.view.dtToQDate(self.sample[6]))
                self.aerobicTable = self.resultToTable(self.sample[7])
                self.anaerobicTable = self.resultToTable(self.sample[8])
                self.comment.setText(self.sample[9])
                self.initTables()
                self.save.setEnabled(True)
                self.clear.setEnabled(True)
                self.preliminary.setEnabled(False)
                self.perio.setEnabled(False)
                self.directSmears.setEnabled(False)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleSavePressed(self):
        try:
            aerobic = self.tableToResult(self.aerobicTable)
            anaerobic = self.tableToResult(self.anaerobicTable)
            if self.model.addCultureResult(
                int(self.sampleID.text()),
                self.chartNumber.text(),
                self.view.entries[self.clinician.currentText()]['db'],
                self.sample[2],
                self.sample[3],
                self.dateReported.date(),
                aerobic,
                anaerobic,
                self.comment.toPlainText()
            ):
                self.handleSearchPressed()
                self.save.setEnabled(False)
                self.clear.setEnabled(False)
                self.preliminary.setEnabled(True)
                self.perio.setEnabled(True)
                self.directSmears.setEnabled(True)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePrintD(self):
        try:
            path = str(Path().resolve())+r'\COMBDb\templates\culture_smear_template.docx'
            dest = self.app.tempify(path)
            document = MailMerge(path)
            document.merge(
                saID=f'{self.sampleID.text()[0:2]}-{self.sampleID.text()[2:6]}',
                chID=self.chartNumber.text(),
                pName=f'{self.sample[3]}, {self.sample[2]}',
                colDate=self.view.fSlashDate(self.sample[4]),
                recDate=self.view.fSlashDate(self.recDate.date())
            )
            document.write(dest)
            self.app.convertAndPrint(dest)
        except Exception as e:
            self.app.showErrorScreen(e)
    
    def handlePrintP(self):
        try:
            path = str(Path().resolve())+r'\COMBDb\templates\culture_prelim_template.docx'
            dest = self.view.tempify(path)
            document = MailMerge(path)
            document.merge(
                saID=formatSaID(self.saID.text()),
                chID=formatChID(self.chID.text()),
                type=self.type.currentText(),
                colDate=formatDate(self.colDate.date()),
                recDate=formatDate(self.recDate.date()),
                repDate=formatDate(self.repDate.date()),
                clName=self.clName.currentText(),
                paName=f'{self.sample[3]}, {self.sample[2]}',
                teName=f'{self.model.tech[1][0]}.{self.model.tech[2][0]}.{self.model.tech[3][0]}.',
                smear=self.smear.toPlainText(),
                comms=self.comms.toPlainText(),
                notes=self.notes.toPlainText()
            )
            document.write(dest)
            context = {
                'headers': ['Aerobic Bacteria']+self.aerobicTable[0][1:],
                'servers': []
            }
            for i in range(1, len(self.aerobicTable)):
                context['servers'].append(self.aerobicTable[i])
            document = DocxTemplate(dest)
            document.render(context)
            document.save(dest)
            self.view.convertAndPrint(dest)
        except Exception as e:
            self.view.showErrorScreen(e)

    def handlePerioPressed(self):
        try:
            template = str(Path().resolve())+r'\COMBDb\templates\culture_results_template.docx'
            dst = self.view.tempify(template)
            document = MailMerge(template)
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
            document.write(dst)
            #aerobic
            context = {
                'headers1' : ['Aerobic Bacteria']+self.aerobicTable[0][1:],
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
        except Exception as e:
            self.view.showErrorScreen(e)

    def handleBackPressed(self):
        self.view.showResultEntryNav()

    def handleReturnToMainMenuPressed(self):
        self.view.showAdminHomeScreen()

class IndexedComboBox(QComboBox):
    def __init__(self, row, column, form, kind):
        super(IndexedComboBox, self).__init__()
        self.row = row
        self.column = column
        self.form = form
        self.kind = kind
        self.currentIndexChanged.connect(self.handleCurrentIndexChanged)

    def handleCurrentIndexChanged(self):
        self.form.updateTable(self.kind, self.row, self.column)