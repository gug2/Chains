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
import random as rand

def clearCellsColor(sheet):
    startCell = getStartCell(sheet)
    
    for row in sheet.iter_rows(startCell.row, sheet.max_row, startCell.column, sheet.max_column):
        for cell in row:
            cell.fill = pyxl.styles.PatternFill(start_color=None)

def getStartCell(sheet):
    for i in range(1, sheet.max_row): #because row starts with 1
        for j in range(0, sheet.max_column):
            if sheet[i][j].value != None:
                return sheet[i][j]
    
    return None

def toMatrix(sheet):
    startCell = getStartCell(sheet)
    
    matrix = []
    for row in sheet.iter_rows(startCell.row, sheet.max_row, startCell.column, sheet.max_column):
        matrixRow = []
        for cell in row:
            if cell.value != None:
                matrixRow.append(cell.value)
        matrix.append(matrixRow)

    return matrix

def markChain(sheet, chain, matrix):
    global markOffset
    startCell = getStartCell(sheet)
    
    r = rand.randint(0, 255)
    g = rand.randint(0, 255)
    b = rand.randint(0, 255)
    randColor = hex(255 << 24 | r << 16 | g << 8 | b)[2::]

    sheetMaxCol = sheet.max_column
    for tupl in chain:
        i, j, elem = tupl
        # invers indexes
        i = len(matrix)-1 - i
        i += 1 # because rows starts with 1
        # offset for start cell
        i += startCell.row-1
        j += startCell.column-1
        #print(randColor)
        sheet[i][j].fill = pyxl.styles.PatternFill(start_color=randColor, fill_type='solid')
        result = sheet.cell(row=i, column=sheetMaxCol+1)
        result.value = sheet[i][j].value
        
# == OPEN PY XL -- END OF MODULE == #

class Main(QMainWindow, gui.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        workbook = pyxl.open('sample.xlsx')
        sheet = workbook.active
        clearCellsColor(sheet)
        
        matrix = [[1,2,3,4,5], [6,7,8,9,10], [5,2,3,5,7]]
        self.table1 = createTable(self, matrix)
        
        chains.matrix = toMatrix(sheet)
        chains.matrix = chains.matrix[::-1]
        self.table2 = createTable(self, chains.matrix)
        self.table3 = createTable(self, matrix)
        
        for i in range(len(chains.matrix)):
            for j in range(len(chains.matrix[i])):
                chainsarray = chains.chainsFor(chains.matrix[i][j], i, j)
                if chainsarray:
                    for i2 in range(len(chainsarray)):
                        for j2 in range(len(chainsarray[i2])):
                            print(chainsarray[i2][j2])

                            # Separate each chain to new sheet for debug
                            #newSheet = workbook.create_sheet('newsheet' + str(i2*len(chainsarray[i2])+j2))
                            #copy
                            #for row in sheet:
                            #    for cell in row:
                            #        newSheet[cell.coordinate].value = cell.value                
                            # mark cells by colors
                            markChain(sheet, chainsarray[i2][j2], chains.matrix)
                            
        
        workbook.save('sample2.xlsx')
        
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
