import xlsxwriter

# Running a sample program
workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'HelloWorld')
workbook.close()

