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

#def updateTable(self, widgetIndex, matrix, posRow, posColumn, rowSpan=1, colSpan=1):
#    widget = self.grid.takeAt(widgetIndex).widget()
#    if widget:
#        widget.setParent(None)
#        table = createTable(self, matrix)
#        self.grid.addWidget(table, posRow, posColumn, rowSpan, colSpan)

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
    self.grid.addWidget(self.label_1, 0, 0)
    self.grid.addWidget(self.table1, 1, 0)
    self.grid.addWidget(self.detectChainsAction, 2, 0)
    self.grid.addWidget(self.label_2, 3, 0)
    self.grid.addWidget(self.table2, 4, 0)
    self.grid.addWidget(self.label_3, 0, 1)
    self.grid.addWidget(self.table3, 1, 1, 4, 1)

# == PY QT -- END OF MODULE == #

class Main(QMainWindow, gui.Ui_MainWindow):
    def openImportMenu(self):
        self.setImportPath()
        
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

        if not self.exportPath:
            # default export path
            self.setExportPath(self.importPath[:self.importPath.rfind('/')])
        
        print('Открыт файл ', self.importPath)     
    
    def setExportPath(self, newPath=None):
        if not newPath:
            self.exportPath = QFileDialog.getExistingDirectory(self)
        else:
            self.exportPath = newPath
        self.setWindowTitle('Последовательности - ' + self.exportPath)

    def setImportPath(self):
        self.importPath = QFileDialog.getOpenFileName(self, 'Открытие', '', 'Таблицы Excel (*.xlsx;*.xlsm;*.xltx;*.xltm)')[0]
        self.label_2.setText('Введенные значения\n' + self.importPath)
    
    def detectChains(self):
        if not self.workbook or not chains.matrix or not self.exportPath:
            return

        print('Последовательности:')
        
        chainsForMark = []
        for i in range(len(chains.matrix)):
            for j in range(len(chains.matrix[i])):
                if chains.matrix[i][j] == '':
                    continue # skip if element is empty
                
                elemChains = chains.chainsFor(chains.matrix[i][j], i, j)
                if elemChains:
                    chainsForMark.append(elemChains)

        startColForChains = max(len(chains.matrix[i]) for i in range(len(chains.matrix)))
        i = 1
        for chain in chainsForMark:
            for row in chain:
                for elem in row:
                    print(elem)
                    pyxlimpl.markChain(self.sheet, elem, startColForChains+i, chains.matrix)
                    i += 1

        print('Последовательности окончены')
        pyxlimpl.saveWorkbook(self.workbook, self.exportPath + '/sample2.xlsx')
        print('Результат сохранен в ' + self.exportPath + '/sample2.xlsx')

        chainsTable = []
        minRow = pyxlimpl.getStartCell(self.sheet).row
        minCol = pyxlimpl.getStartCell(self.sheet).column + max(len(chains.matrix[i]) for i in range(len(chains.matrix)))
        for row in self.sheet.iter_rows(minRow, self.sheet.max_row, minCol, self.sheet.max_column):
            chainsRow = []
            for cell in row:
                if cell.value and type(cell.value) == int:
                    chainsRow.append(cell.value)
                else:
                    chainsRow.append('')
            chainsTable.append(chainsRow)

        # update layouts
        self.table3Matrix = chainsTable
        updateLayout(self)
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        self.workbook = None
        self.importPath = None
        self.exportPath = None
        self.importAction.triggered.connect(self.openImportMenu)
        self.exportPathAction.triggered.connect(self.setExportPath)

        # 1 table
        self.label_1 = QLabel(self)
        self.label_1.setText('Предсказанные значения')

        # detect chains action
        self.detectChainsAction = QPushButton(self)
        self.detectChainsAction.setText('Последовательности')
        self.detectChainsAction.clicked.connect(self.detectChains)

        # 2 table
        self.label_2 = QLabel(self)
        self.label_2.setText('Введенные значения')
        
        # 3 table
        self.label_3 = QLabel(self)
        self.label_3.setText('Последовательности')
        
        self.grid = QGridLayout(self.centralwidget)
        self.table1Matrix = [[]]
        self.table2Matrix = [[]]
        self.table3Matrix = [[]]
        updateLayout(self)
     
app = QApplication(sys.argv)
window = Main()
window.show()
sys.exit(app.exec())
