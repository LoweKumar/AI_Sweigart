import openpyxl
# from openpyxl import open

def readingValue():
    wb = openpyxl.open('example.xlsx')
    print(wb.get_sheet_names())
    sheet = wb.get_sheet_by_name('Sheet1')
    print(tuple(sheet['A1':'C3']))
    # print(sheet.get_highest_column())
    for i in range(1, 8, 2):
        print(i, sheet.cell(row=i, column=1).value)
    
    for rowOfCellObjects in sheet['A1':'C3']: 
        for cellObj in rowOfCellObjects:
            print(cellObj.coordinate, cellObj.value)
    print('--- END OF ROW ---')

    # sheetActive = wb.get_active_sheet()
    # print(sheetActive.columns[1])
    # for cellObjValue in sheetActive.columns[1]:
    #     print(cellObjValue.value)



if __name__ == '__main__':
    readingValue()