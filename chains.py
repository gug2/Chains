matrix = [
    [6,9,1,4,7,21],
    [11,15,9,2,5,25,26],
    [12,44,43,42,21,7,20],
    [8,28,12,16,56,14,3],
    [1,2,8,6,17,18,15],
    [44,24,43,12,27,10,8,77,1],
    [18,9,17,16,2,14,6],
    [3,25,9,47,6,8,13,99,53],
    [10,46,68,23,41,11,7,88,69]
]
matrix = matrix[::-1]

def printTable(matrix):
    for row in matrix:
        print(row)

def findElement(row, which):
    if which in row:
        return True

    return False

def findNotDigits(row):
    rowStr = ''
    for element in row:
        rowStr += str(element)
    
    notDigits = []

    # 0 - 9
    for digit in range(1, 10):
        if str(digit) not in rowStr:
            notDigits.append(digit)
    
    return notDigits

def chainsFor(element, rowIndex):
    for i in range(-32, 0):
        chainFor(element, element+i, rowIndex)
    for i in range(1, 33):
        chainFor(element, element+i, rowIndex)
    
    #chainFor(element, element+1, rowIndex)
    #chainFor(element, element-1, rowIndex)
    #chainFor(element, element+3, rowIndex)
    #chainFor(element, element-3, rowIndex)

def chainFor(element, nextElement, rowIndex):
    chainList = []
    chain = []

    currentLc = element
    nextLc = nextElement
    
    for step in range(1, len(matrix) // 2 + 1):
        # new step - clear current and next elements
        currentLc = element
        nextLc = nextElement
        chain.append( (rowIndex, currentLc) )
        flagValid = 1
        for i in range(rowIndex, len(matrix), step):
            if i == rowIndex: continue
            if findElement(matrix[i], nextLc):
                chain.append( (i, nextLc) )
                nextDelta = nextLc - currentLc
                currentLc = nextLc
                nextLc += nextDelta
            else:
                #print('not finded', matrix[i], nextLc)
                if len(chain) > 2 and i + step <= len(matrix) + 1:
                    flagValid = 0
                    # т.к закончилась раньше конца таблицы :(
                    print('* последовательность недействительна', chain)
                break
        # current step end - clear local chain and add to list
        if flagValid == 1 and len(chain) > 2:
            chainList.append(chain)
        chain = []
    
    
    if len(chainList) > 0:
        print('* List of chains', nextElement-element, 'for', element, 'in row', rowIndex)
        printTable(chainList)

#print("==Дано==")
#printTable(matrix)
#print("==Последовательности где нет цифр в строке==")
#for row in matrix:
#    print(findNotDigits(row))
#for i in range(len(matrix)):
#    for j in range(len(matrix[i])):
#        #print("==Последовательности для числа", matrix[i][j], 'в строке', i, "==")
#        chainsFor(matrix[i][j], i)