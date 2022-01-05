import sys
import gui
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import *
import chains
import pyxlimpl

# == PY QT -- MODULE == #
def createTable(self, matrix):
    table = QTableWidget(self)
    table.setRowCount(len(matrix))
    # get max columns in matrix
    maxColumns = max(len(matrix[i]) for i in range(len(matrix)))
    table.setColumnCount(maxColumns)
    
    for i in range(table.rowCount()):
        for j in range(table.columnCount()):
            item = QTableWidgetItem(str(matrix[i][j]))
            item.setBackground(QBrush(QColor(255, 255, 255)))
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(i, j, item)
    
    table.resizeRowsToContents()
    table.resizeColumnsToContents()
    
    return table

def updateLayout(self):
    # remove all widgets
    for i in reversed(range(self.grid.count())):
        self.grid.itemAt(i).widget().setParent(None)
    
    initWidgets(self)

def initWidgets(self):
    # 1 table
    self.table1 = createTable(self, self.table1Matrix)
    
    # 2 table
    self.table2 = createTable(self, self.table2Matrix)
    
    # 3 table
    self.table3 = createTable(self, self.table3Matrix)
    
    #grid
    self.grid.addWidget(self.label_1, 0, 0, 1, 2)
    self.grid.addWidget(self.table1, 1, 0, 2, 2)
    self.grid.addWidget(self.detectChainsAction, 3, 0)
    self.grid.addWidget(self.continueRowsSetter, 3, 1)
    self.grid.addWidget(self.label_2, 4, 0, 1, 2)
    self.grid.addWidget(self.table2, 5, 0, 1, 2)
    self.grid.addWidget(self.label_3, 0, 2, 1, 2)
    self.grid.addWidget(self.open_3, 1, 2, 1, 2)
    self.grid.addWidget(self.table3, 2, 2, 4, 2)
    
# == PY QT -- END OF MODULE == #

class Main(QMainWindow, gui.Ui_MainWindow):
    def importWorkbook(self):
        self.importPath = self.getImportPath()
        
        if not self.importPath:
            return
        
        self.workbook = pyxlimpl.getWorkbook(self.importPath)
        self.sheet = pyxlimpl.getSheet(self.workbook)

        chains.matrix = pyxlimpl.toMatrix(self.sheet)

        # update layouts
        self.table2Matrix = chains.matrix
        updateLayout(self)

        # inverse rows for chain detection
        chains.matrix = chains.matrix[::-1]

        if not self.exportDirectory:
            # default export directory - import directory
            self.setExportDirectory(self.getDirectoryFromPath(self.importPath))
        
        print('Открыт файл', self.importPath)

    def getImportPath(self):
        path = QFileDialog.getOpenFileName(self, 'Открытие', '', 'Таблицы Excel (*.xlsx;*.xlsm;*.xltx;*.xltm)')[0]

        if not path:
            return
        
        self.label_2.setText('Введенные значения\n' + path)
        return path

    def setExportDirectory(self, directory=None):
        exportDir = None
        
        if directory: exportDir = directory
        else: exportDir = QFileDialog.getExistingDirectory(self) + '/'

        if not exportDir:
            return
        
        self.exportDirectory = exportDir
        self.setWindowTitle('Последовательности - ' + exportDir)

    def getDirectoryFromPath(self, path):
        if path.rfind('.') == -1: return # non path value
        
        return path[:path.rfind('/')+1]

    def getFileNameFromPath(self, path):
        if path.rfind('.') == -1: return # non path value
        
        return ( path[path.rfind('/')+1:path.rfind('.')], path[path.rfind('.'):len(path)] )

    def save(self, currentFilePath, overwrite=False):
        import os.path        
        filename, extention = self.getFileNameFromPath(currentFilePath)
        
        if not filename or not extention:
            return

        tempFile = None
        if overwrite:
            tempFile = self.lastSavePath
        else:
            i = 1
            tempFile = self.exportDirectory + filename + str(i) + extention
            while os.path.exists(tempFile):
                i += 1
                tempFile = self.exportDirectory + filename + str(i) + extention
        
        pyxlimpl.saveWorkbook(self.workbook, tempFile)
        self.label_3.setText('Последовательности\n' + tempFile)
        print('Результат сохранен в', tempFile)

        return tempFile

    def openResultWorkbook(self):
        if not self.lastSavePath:
            return
        
        self.openFile(self.lastSavePath)
    
    def openFile(self, path):
        import os
        if not os.path.exists(path):
            return
        
        os.startfile(path)
        print('Открытие', path)
    
    def detectChains(self):
        if not self.workbook or not self.sheet or not chains.matrix or not self.exportDirectory:
            return
        
        self.continueRows = self.continueRowsSetter.value()
        
        print('Последовательности:')
        # clear cells color
        pyxlimpl.clearCellsColor(self.sheet)

        
        # detect chains for all elements in given table
        chainsForMark = []
        for i in range(len(chains.matrix)):
            for j in range(len(chains.matrix[i])):
                if chains.matrix[i][j] == '':
                    continue # skip element if it's empty
                
                elemChains = chains.chainsFor(chains.matrix[i][j], i, j)
                if elemChains:
                    chainsForMark.append(elemChains)


        # continue detected chains
        continueChainsTable = []
        # mark all chain in result table
        startColForChains = max(len(chains.matrix[i]) for i in range(len(chains.matrix)))
        i = 1
        for chain in chainsForMark:
            for row in chain:
                for elem in row:
                    print(elem)
                    continueChainsTable.append( pyxlimpl.continueChain(self.sheet, elem, startColForChains+i, self.continueRows, chains.matrix) )
                    pyxlimpl.markChain(self.sheet, elem, startColForChains+i, self.continueRows, chains.matrix)
                    i += 1
        
        print('Последовательности окончены')

        # fill chainsTable from excel sheet
        chainsTable = []
        minRow = pyxlimpl.getStartCell(self.sheet).row
        minCol = pyxlimpl.getStartCell(self.sheet).column + max(len(chains.matrix[i]) for i in range(len(chains.matrix)))
        for row in self.sheet.iter_rows(minRow, self.sheet.max_row, minCol, self.sheet.max_column):
            chainsRow = []
            for cell in row:
                if cell.row == self.continueRows:
                    cell.fill = pyxlimpl.pyxl.styles.PatternFill(start_color='FF0000', fill_type='solid')
                if cell and type(cell.value) == int:
                    chainsRow.append(cell.value)
                else:
                    chainsRow.append('')
            chainsTable.append(chainsRow)
        
        # save workbook
        self.lastSavePath = self.save(self.importPath, self.lastSavePath and self.shouldOverwrite)

        # update layouts
        # transponse table 1
        continueChainsTableTransponse = []
        for j in range(len(continueChainsTable[0])):
            continueChainsTableTransponse.append([ continueChainsTable[i][j] for i in range(len(continueChainsTable)) ])
        # inverse table 1
        continueChainsTableTransponse = continueChainsTableTransponse[::-1]
        
        self.table1Matrix = continueChainsTableTransponse
        self.table3Matrix = chainsTable
        updateLayout(self)

        # return to default given excel sheet
        self.workbook = pyxlimpl.getWorkbook(self.importPath)
        self.sheet = pyxlimpl.getSheet(self.workbook)

    def setOverwrite(self):
        self.shouldOverwrite = not self.shouldOverwrite
        s = 'Да'
        if not self.shouldOverwrite:
            s = 'Нет'
        self.overwriteAction.setText('Перезаписывать файл - ' + s)
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.continueRows = 0
        
        self.lastSavePath = None
        self.workbook = None
        self.sheet = None
        self.importPath = None
        self.exportDirectory = None
        self.importAction.triggered.connect(self.importWorkbook)
        self.exportDirectoryAction.triggered.connect(self.setExportDirectory)
        # overwrite settings
        self.shouldOverwrite = False
        self.setOverwrite()
        self.overwriteAction.triggered.connect(self.setOverwrite)
        
        # 1 table
        self.label_1 = QLabel(self)
        self.label_1.setText('Предсказанные значения')

        # detect chains action
        self.detectChainsAction = QPushButton(self)
        self.detectChainsAction.setText('Последовательности')
        self.detectChainsAction.clicked.connect(self.detectChains)

        # continue rows setter
        self.continueRowsSetter = QSpinBox(self)
        self.continueRowsSetter.setSuffix(' строк')
        self.continueRowsSetter.setSingleStep(1)
        self.continueRowsSetter.setRange(1, 999)

        # 2 table
        self.label_2 = QLabel(self)
        self.label_2.setText('Введенные значения')
        
        # 3 table
        self.label_3 = QLabel(self)
        self.label_3.setText('Последовательности')
        
        # open button
        self.open_3 = QPushButton(self)
        self.open_3.setText('Открыть таблицу')
        self.open_3.clicked.connect(self.openResultWorkbook)
        
        self.grid = QGridLayout(self.centralwidget)
        self.table1Matrix = [[]]
        self.table2Matrix = [[]]
        self.table3Matrix = [[]]
        updateLayout(self)
     
app = QApplication(sys.argv)
window = Main()
window.show()
sys.exit(app.exec())
