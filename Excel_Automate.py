


import datetime, time

import pandas as pd
import xlwings as xw
import time


# python.exe -m pip install --upgrade pip
# pip install pandas
# pip install xlwings



pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", None)



def time_():
    time_ = time.strftime("%H:%M:%S", time.localtime())
    return (time_)



# def active_Workbook_name():
#     try:
#         wb = xw.apps.active.books.active
#         print(wb.name)
#         return wb.name
#     except  Exception as e:  # This catches xlwings-specific errors
#         print("time",time_(),"opne is Excel.",e)
#         time.sleep(1)
#
#         # active_Workbook_name()
#         return "opne Excel"

# active_Workbook_name()



import xlwings as xw
def active_Workbook_name():
    if xw.apps:
       wb = xw.apps.active.books.active
       print(wb.name)
       return wb.name

    else:
         print("No active Excel application found.")

#
# active_Workbook_name()


def sheet_names_list():
    wb = xw.apps.active.books.active
    sheet_names_list = [sheet.name for sheet in wb.sheets]
    # Print the list of sheet names
    # print("List of sheet names:", sheet_names_list)
    return sheet_names_list


# sheet_names_list()
#
# print(len(sheet_names_list())-1)
# print((1) - 1)


def active_sheet_names():
    wb = xw.apps.active.books.active
    active_sheet_name = wb.sheets.active.name
    # print("active_sheet_names:", active_sheet_name)
    return active_sheet_name


# active_sheet_names()


def active_tabal_selet(range):
    wb = xw.apps.active.books.active
    active_sheet = wb.sheets.active
    range_to_name = active_sheet.range(range).value
    # Convert the range data to a pandas DataFrame for better manipulation
    df = pd.DataFrame(range_to_name[1:], columns=range_to_name[0])  # Using first row as headers
    # print(df)
    return df


# df = active_tabal_selet(range="A1:Y100")


def pandas_convert_dict(df):
    # Create header_dict with actual data
    header_dict = [{col: df[col].tolist() for col in df.columns}]
    # print(header_dict)
    # Convert header_dict back to a DataFrame
    df_from_dict = pd.DataFrame(header_dict[0])  # Extract dictionary from list and convert to DataFrame
    # print("Converted DataFrame from header_dict:",df_from_dict)
    return header_dict
# pandas_convert_dict(df=df)


def sheets_tabal_selet(sheet, range):
    wb = xw.books.active  # This fetches the active workbook
    sheet = wb.sheets[sheet]  # Access the sheet named 'Sheet1'
    range_to_name = sheet.range(range).value
    # Convert the range data to a pandas DataFrame for better manipulation
    df = pd.DataFrame(range_to_name[1:], columns=range_to_name[0])  # Using first row as headers
    # print(df)
    return df


def range_to_row(vlookup_sheets, filtered_row):
    wb = xw.books.active  # This fetches the active workbook
    sheet = wb.sheets[vlookup_sheets]  # Access the sheet named 'Sheet1'

    range_to_row = sheet.range(filtered_row).value
    range_to_row_Remove_None = list(filter(None, dict.fromkeys(range_to_row )))
    print(range_to_row_Remove_None)
    return  range_to_row_Remove_None




def df_vlookup(df,vlookup_sheets, filtered_columns, filtered_row):
    global  range_to_columns_Remove_None
    wb = xw.books.active  # This fetches the active workbook
    sheet = wb.sheets[vlookup_sheets]  # Access the sheet named 'Sheet1'
    range_to_columns = sheet.range(filtered_columns).value

    # Remove None values and duplicates from the list
    range_to_columns_Remove_None = list(filter(None, dict.fromkeys(range_to_columns)))
    # print(range_to_columns_Remove_None)
    range_to_row = sheet.range(filtered_row).value
    range_to_row_Remove_None = list(filter(None, dict.fromkeys(range_to_row )))
    # print(range_to_row_Remove_None)
    df_filtered = df[df[range_to_row_Remove_None[0]].isin(range_to_row_Remove_None)][range_to_columns_Remove_None]
    print(df_filtered)
    # breakpoint()     converters  excel








    return df_filtered
#
def time_():
    time_ = time.strftime("%H:%M:%S", time.localtime())
    return (time_)


def final_output_cell(Sheet_names,df,range):
    wb = xw.books.active  # This fetches the active workbook
    sheet = wb.sheets[Sheet_names]

    # # set index column name
    df.set_index(df.columns[0], inplace=True)
    sheet.range(range).value = df
    # print('final_output_cell', time_())
    #

# Function to convert column numbers to letters
def number_to_excel_column(n):
    letters = ""
    while n > 0:
        n -= 1
        letters = chr(n % 26 + ord('A')) + letters
        n //= 26
    return letters




def column_letters_list():
    # Generate column letters for numbers 1 to 16384
    column_numbers = range(1, 16385)  # Numbers from 1 to 16384
    column_letters = [number_to_excel_column(num) for num in column_numbers]

    # Print the result (first 20 and last column)
    # print("First 20 Column Letters:", column_letters[:20])  # First 20 columns
    # print("Last Column Letter:", column_letters[-1])  # Last column (XFD)
    # breakpoint()

    return   column_letters
# column_letters_list()




# Generate column letters for numbers 1 to 16384
# column_numbers = range(1, 16385)  # Numbers from 1 to 16384
# column_letters = [number_to_excel_column(num) for num in column_numbers]




# Print the result (first 20 and last column)
# print("First 20 Column Letters:", column_letters[:20])  # First 20 columns
# print("Last Column Letter:", column_letters[-1])  # Last column (XFD)

# breakpoint()

# column_numbers = range(1, 16385)  # Numbers from 1 to 16384
# column_letters = [number_to_excel_column(num) for num in column_numbers]
# index = column_letters.index('A')


# print(column_letters)
# breakpoint()
def split_Column(text):
    import re
    # text = "sheet2!AAAAA110"
    match = re.match(r'([a-zA-Z0-9_]+)!(\w+)(\d+)', text)
    if match:
        sheet_name = match.group(1)
        cell_column = match.group(2)
        cell_row = match.group(3)
        # print("Sheet Name:", sheet_name)  # Output: sheet2
        # print("Column:", cell_column)  # Output: AA
        # print("Row:", cell_row)  # Output: 110
        return cell_column


split_Column(text="sheet2!AAAAA110")


def split_cell_column(cell_column):
    import re
    text = cell_column
    # Extract the column (AA) and row (111)
    column = ''.join(filter(str.isalpha, text))  # Output: AA
    row = ''.join(filter(str.isdigit, text))  # Output: 111
    # print("Column:", column)  # Output: AA
    # print("Row:", row)
    return column, row













# split_cell_column(cell_column=split_Column(text="sheet2!AAAAA110"))
#
# split_cell_column(cell_column='AAAAA110')
#





# print(split_cell_column[0],split_cell_column[1])
#


# print(column_letters[:20])
#
# print(index)
# print(column_letters[index + 1 ] )
#


# my_list = [10, 20, 30, 20, 40]
# item_to_find = 20
# indices = [index for index, value in enumerate(my_list) if value == item_to_find]
# print("20 के इंडेसेस:", indices)
# # index = column_letters.index('c')


# print(column_indices)

# df = sheets_tabal_selet(sheet='Sheet1', range="A1:Y30")
#
# data_vlookup = df_vlookup(df=df,vlookup_sheets='Sheet2', filtered_columns='A1:Y1', filtered_row='B1:B15')
#
# print(data_vlookup.to_dict())
# final_output_cell(Sheet_names='Sheet3',df=data_vlookup,range="A18")

# #


    # wb = xw.books.active  # This fetches the active workbook
    # sheet = wb.sheets['Sheet3']
    #
    # # # set index column name
    # data_vlookup.set_index(data_vlookup.columns[0], inplace=True)
    # sheet.range('A1').value = data_vlookup




#  Install packages from requirements.txt: Run this command in your terminal or command prompt where the file is located.

# pip install -r requirements.txt


# Generate the requirements.txt File
# pip freeze > requirements.txt

