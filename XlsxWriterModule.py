import xlsxwriter

# Running a sample program
# workbook = xlsxwriter.Workbook('hello.xlsx')
# worksheet = workbook.add_worksheet()
# worksheet.write('A1', 'HelloWorld')
# workbook.close()

# Create a workbook and add a worksheet.
# https://xlsxwriter.readthedocs.io/tutorial01.html

# workbook = xlsxwriter.Workbook('Expenses01.xlsx')
# worksheet = workbook.add_worksheet()

# # Some data we want to write to the worksheet.
# expenses = (
#     ['Rent', 1000],
#     ['Gas',   100],
#     ['Food',  300],
#     ['Gym',    50],
# )

# # Start from the first cell. Rows and columns are zero indexed.
# row = 0
# col = 0

# # Iterate over the data and write it out row by row.
# for item, cost in (expenses):
#     worksheet.write(row, col, item)
#     worksheet.write(row, col+1, cost)
#     row+=1

# # Write a total using a formula.
# worksheet.write(row, 0, 'Total')
# worksheet.write(row, 1, '=SUM(B1:B4)')

# workbook.close()

# Adding formatting to the XLSX File
# # https://xlsxwriter.readthedocs.io/tutorial02.html
# workbook = xlsxwriter.Workbook('Expenses02.xlsx')
# worksheet = workbook.add_worksheet()

# # Add a bold format to use to highlight cells
# bold = workbook.add_format({'bold':True})
# # Add a number format for cells with money.
# money = workbook.add_format({'num_format': '$#,##0'})

# # Write some data headers
# worksheet.write('A1', 'Item', bold)
# worksheet.write('B1', 'Cost', bold)

# # Some data we want to write to the worksheet.
# expenses = (
#      ['Rent', 1000],
#      ['Gas',   100],
#      ['Food',  300],
#      ['Gym',    50],
#  )

#  # Start from the first cell below the headers.
# row = 1
# col = 0

# # Iterate over the data and write it out row by row.
# for item, cost in (expenses):
#     worksheet.write(row, col, item)
#     worksheet.write(row, col+1 ,cost, money)
#     row+=1

# # Write a total using a formula.
# worksheet.write(row, 0, 'Total',       bold)
# worksheet.write(row, 1, '=SUM(B2:B5)', money)

# workbook.close()

# ==============================================================================

# Writing different types of data to the XLSX File
# https://xlsxwriter.readthedocs.io/tutorial03.html

# from datetime import datetime
# # Create a workbook and add a worksheet.
# workbook = xlsxwriter.Workbook('Expenses03.xlsx')
# worksheet = workbook.add_worksheet()

#  # Add a bold format to use to highlight cells.
# bold = workbook.add_format({'bold': 1})

#  # Add a number format for cells with money.
# money_format = workbook.add_format({'num_format': '$#,##0'})

#  # Add an Excel date format.
# date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

# # Adjust the column width.
# worksheet.set_column(1, 1, 15)

#  # Write some data headers.
# worksheet.write('A1', 'Item', bold)
# worksheet.write('B1', 'Date', bold)
# worksheet.write('C1', 'Cost', bold)

# # Some data we want to write to the worksheet.
# expenses = (
#      ['Rent', '2013-01-13', 1000],
#      ['Gas',  '2013-01-14',  100],
#      ['Food', '2013-01-16',  300],
#      ['Gym',  '2013-01-20',   50],
#  )

#  # Start from the first cell below the headers.
# row = 1
# col = 0

# for item, date_str, cost in (expenses):
#      # Convert the date string into a datetime object.
#     date = datetime.strptime(date_str, "%Y-%m-%d")
#     worksheet.write_string(row, col, item)
#     worksheet.write_datetime(row, col + 1, date, date_format)
#     worksheet.write_number(row, col + 2, cost, money_format)
#     row+=1

# # Write a total using a formula.
# worksheet.write(row, 0, 'Total', bold)
# worksheet.write(row, 2, '=SUM(C2:C5)', money_format)

# workbook.close()

# *******************************************************************************
# For a list of lists structure you would use two loop levels:
# https://xlsxwriter.readthedocs.io/working_with_data.html

# workbook = xlsxwriter.Workbook('write_list.xlsx')
# worksheet = workbook.add_worksheet()

# my_list = [[1, 1, 1, 1, 1],
#            [2, 2, 2, 2, 1],
#            [3, 3, 3, 3, 1],
#            [4, 4, 4, 4, 1],
#            [5, 5, 5, 5, 1]]

# for row_num, row_data in enumerate(my_list):
#     for col_num, col_data in enumerate(row_data):
#         worksheet.write(row_num, col_num, col_data)

# workbook.close()

# Writing dicts of data
# workbook = xlsxwriter.Workbook('write_dict.xlsx')
# worksheet = workbook.add_worksheet()

# my_dict = {'Bob': [10, 11, 12],
#            'Ann': [20, 21, 22],
#            'May': [30, 31, 32]}

# col_num = 0
# for key, value in my_dict.items():
#     worksheet.write(0, col_num, key)
#     worksheet.write_column(1, col_num, value)
#     col_num += 1

# workbook.close()

# Writing dataframes
import pandas as pd

# Create a Pandas dataframe from the data.
df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.close()







