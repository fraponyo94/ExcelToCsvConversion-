import sys
import os
import xlrd
import openpyxl
import unicodecsv
from glob import glob


# Excel files directory
source="/home/fred/apache-nifi/input/"

# Converted Csv files directory
dest='/home/fred/apache-nifi/output/'

os.chdir(source)



# Handle  excel file conversion to csv
def excelToCsvConvertor (excel_filename):
    excel_filename_extension = excel_filename.rsplit('/', 1)[-1].rsplit('.', 1)[-1]
    print(excel_filename_extension)
    
    try:
        if excel_filename_extension in ("xls","XLS"):
            # load the workbook for xls 
            work_book = xlrd.open_workbook(excel_filename)

            # Check the number of sheets in the workbook.
            no_of_sheets = work_book.nsheets

        elif excel_filename_extension in ("xlsx","XLSX"):
            # load the workbook for xlsx 
            work_book = openpyxl.load_workbook(excel_filename)

            # Check the number of sheets in the workbook .
            no_of_sheets = len(work_book.get_sheet_names())
            
            # Get sheet names in th workbook
            sheet_names = work_book.get_sheet_names()

        # # Check the number of sheets in the workbook .
        # no_of_sheets = work_book_xls.nsheets
        print("Number of sheets {}".format(no_of_sheets))
      

        # Loop through the all the sheets.
        for sheet_number in range(no_of_sheets):
           
            if excel_filename_extension in ("xls","XLS"):
                try:                   
                
                    # Open the sheet by index for xls sheets
                    sheet = work_book.sheet_by_index(sheet_number)

                    # Open the csv file in binary write mode.
                    with open(dest+"%s.csv" %(sheet.name.replace(" ","")), "wb") as file:
                    
                    # Uses unicodecsv, so it will handle Unicode characters.
                        csv_out = unicodecsv.writer(file, encoding='utf-8')                    

                        header = [cell.value for cell in sheet.row(0)]
                        csv_out.writerow(header)

                        # Loop through the rows of the sheet and write to csv file.
                        for row_number in range(sheet.nrows):
                            csv_out.writerow(sheet.row_values(row_number))

                        # Close the csv file.
                        file.close()

                        print("CSV file created successfully.")

                except:
                    print("Error creating CSV file.")
                    print(sys.exc_info())


            elif excel_filename_extension in ("xlsx","XLSX"):

                try:
                    # Open the sheet by name for xlsx sheets.
                    sheet = work_book[sheet_names[sheet_number]]

                    
                    # Open the csv file in binary write mode.
                    with open(dest+"%s.csv" %(sheet_names[sheet_number].replace(" ","")), "wb") as file:
                    
                    # Uses unicodecsv, so it will handle Unicode characters.
                        csv_out = unicodecsv.writer(file, encoding='utf-8')                    

                        # header = [cell.value for cell in sheet.row(0)]
                        # csv_out.writerow(header)

                        
                        # Loop through the rows of the sheet and write to csv file.
                        for row in sheet.rows:
                            csv_out.writerow(cell.value for cell in row)

                        # Close the csv file.
                        file.close()

                        print("CSV file created successfully.")

                except:
                    print("Error creating CSV file.")
                    print(sys.exc_info())
    except:
        print("Error opening the file.")
        print(sys.exc_info())


if __name__ == '__main__':

   # select only files with extension xls,XLS,xlsx,XLSX
    files = glob('*.xls')
    files.extend(glob('*.XLS'))
    files.extend(glob('*.xlsx'))
    files.extend(glob('*.XLSX'))
       
#     excelToCsvConvertor(sys.argv[1])     
    for file in files:
        excelToCsvConvertor(file)