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
    for j in range(len(row)):
        if row[j] == which:
            return (j, True)
    
    return (-1, False)
    #if which in row:
    #    return True

    #return False

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

def chainsFor(element, rowIndex, colIndex):
    chains = []
    
    chain = []
    #for i in range(-32, 0):
    #    chain = chainFor(element, element+i, rowIndex, colIndex)
    #    if chain:
    #        chains.append(chain)
    #for i in range(1, 33):
    #    chain = chainFor(element, element+i, rowIndex, colIndex)
    #    if chain:
    #        chains.append(chain)
    
    # возрастающие, убывающие +-1
    # четные / нечетные +-2
    # кратные 3,4,...10
    for i in range(1, 10+1):
        if i < 3:
            chain = chainFor(element, element+i, rowIndex, colIndex)
            if chain:
                chains.append(chain)
            chain = chainFor(element, element-i, rowIndex, colIndex)
            if chain:
                chains.append(chain)
        elif element % i == 0:
            chain = chainFor(element, element+i, rowIndex, colIndex)
            if chain:
                chains.append(chain)
            chain = chainFor(element, element-i, rowIndex, colIndex)
            if chain:
                chains.append(chain)
    
    #chainFor(element, element+1, rowIndex)
    #chain = chainFor(element, element-3, rowIndex, colIndex)
    #if chain:
    #    chains.append(chain)
    #chainFor(element, element+3, rowIndex)
    #chainFor(element, element-3, rowIndex)

    return chains

def chainFor(element, nextElement, rowIndex, colIndex):
    chainList = []
    chain = []

    currentLc = element
    nextLc = nextElement
    
    for step in range(1, len(matrix) // 2 + 1):
        # new step - clear current and next elements
        currentLc = element
        nextLc = nextElement
        chain.append( (rowIndex, colIndex, currentLc) )
        flagValid = 1
        curI = rowIndex
        for i in range(rowIndex, len(matrix), step):
            if i == rowIndex: continue
            if findElement(matrix[i], nextLc)[1]:
                j = findElement(matrix[i], nextLc)[0]
            
                chain.append( (i, j, nextLc) )
                curI = i
                nextDelta = nextLc - currentLc
                currentLc = nextLc
                nextLc += nextDelta
            else:
                #print('not finded', matrix[i], nextLc)
                if len(chain) > 2 and curI + step < len(matrix):
                    flagValid = 0
                    # т.к закончилась раньше конца таблицы :(
                    #print('* последовательность недействительна', chain)
                break
        # current step end - clear local chain and add to list
        if flagValid == 1 and len(chain) > 2:
            chainList.append(chain)
        chain = []
    
    
    #if len(chainList) > 0:
    #    print('* List of chains', nextElement-element, 'for', element, 'in row', rowIndex)
    #    printTable(chainList)
        
    return chainList

#print("==Дано==")
#printTable(matrix)
#print("==Последовательности где нет цифр в строке==")
#for row in matrix:
#    print(findNotDigits(row))
#for i in range(len(matrix)):
#    for j in range(len(matrix[i])):
#        #print("==Последовательности для числа", matrix[i][j], 'в строке', i, "==")
#        chainsFor(matrix[i][j], i, j)