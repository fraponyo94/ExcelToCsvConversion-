# xls(before 2007 excel) and xlsx(after ..) conversion to csv
This script was inspired by the limitation witnessed in Apache Nifi when in need of converting excel files contained in a  workbook with more the one worksheet to csv. As per the available capablitiy provided by convertExcelToCSVProcessor,it only supports xlsx. Therefore one has to write his/her own script to handle xls and execute it in Nifi.

This is a python script that converts both xls and xlsx to csv

# Requirements
python => 3

xlrd  library,
openpyxl libray,
unicodecsv libray

if they are not installed in your python environment
>>> pip install [library name]

# Usage
$ python scriptName path_to_excel_file

