
def InputDataToSubject(workbook,sheet,row,value):
    sheet = workbook[str(sheet)]
    num = 7
    while True:
        cell_value = sheet[row+str(num)]
        if cell_value.value is None:
            cell_value.value = value
            print(f"Sum Value {num-3} = {cell_value.value}")
            break
        else:
            num +=1
        

def InputDataToSum(workbook,sheet,row,value):
    sheet = workbook[str(sheet)]
    num = 4
    while True:
        cell_value = sheet[row+str(num)]
        if cell_value.value is None:
            cell_value.value = value
            print(f"Sum Value {num-3} = {cell_value.value}")
            break
        else:
            num +=1

def InputDataToIncome(workbook,sheet,row,value):
    sheet = workbook[str(sheet)]
    num = 3
    sheet[row+str(num)].value = value

def checkRow(workbook,sheet):
    sheet = workbook[str(sheet)]
    num = 7
    cell = sheet['A'+str(num)]
    while True:
        if cell.value is not None:
            cell.value = num-6
            print(num)
            return num
        else:
            num +=1
            print(num)
    
