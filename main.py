import sys
import gui
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.Qt import *
import chains

def createTable(self, matrix):
    table = QTableWidget(self)
    table.setRowCount(len(matrix))
    # get max columns in matrix
    maxColumns = max(len(matrix[i]) for i in range(len(matrix)))
    table.setColumnCount(maxColumns)
    
    for i in range(table.rowCount()):
        for j in range(len(matrix[i])):
            item = QTableWidgetItem(str(matrix[i][j]))
            item.setBackground(QBrush(QColor(255, 255, 255)))
            table.setItem(i, j, item)
    
    table.resizeRowsToContents()
    table.resizeColumnsToContents()
    
    return table

def markChain():
    pass

# == OPEN PY XL -- MODULE == #
import openpyxl as pyxl

def getFirstCell(sheet):
    for i in range(1, sheet.max_row): #because row starts from 1
        for j in range(0, sheet.max_column):
            if sheet[i][j].value != None:
                return sheet[i][j]
    
    return None

def toMatrix(workbookName):
    workbook = pyxl.open(workbookName)
    sheet = workbook.active
    startCell = getFirstCell(sheet)
    
    matrix = []
    for row in sheet.iter_rows(startCell.row, sheet.max_row, startCell.column, sheet.max_column):
        matrixRow = []
        for cell in row:
            if cell.value != None:
                matrixRow.append(cell.value)
        matrix.append(matrixRow)

    return matrix

def markChain(chain):
    for tupl in chain:
        index, elem = tupl

# == OPEN PY XL -- END OF MODULE == #

class Main(QMainWindow, gui.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        matrix = [[1,2,3,4,5], [6,7,8,9,10], [5,2,3,5,7]]
        self.table1 = createTable(self, matrix)
        
        chains.matrix = toMatrix('sample.xlsx')
        chains.matrix = chains.matrix[::-1]
        self.table2 = createTable(self, chains.matrix)
        self.table3 = createTable(self, matrix)

        for i in range(len(chains.matrix)):
            for j in range(len(chains.matrix[i])):
                chains.chainsFor(chains.matrix[i][j], i)
        
        #grid
        self.grid = QGridLayout(self.centralwidget)
        self.grid.addWidget(self.label, 0, 0, 1, 2)
        self.grid.addWidget(self.table1, 1, 0, 1, 2)
        self.grid.addWidget(self.label_2, 2, 0, 1, 2)
        self.grid.addWidget(self.enterButton, 2, 1)
        self.grid.addWidget(self.table2, 3, 0, 1, 2)
        self.grid.addWidget(self.label_3, 0, 2)
        self.grid.addWidget(self.table3, 1, 2, 3, 1)
            
app = QApplication(sys.argv)
window = Main()
window.show()
sys.exit(app.exec())