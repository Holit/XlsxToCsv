import xlrd
import csv  
import sys
import os

if __name__ == '__main__':
    if len(sys.argv) == 1:
        print("illegal argument, try again.")
    elif len(sys.argv[1]) == 0:
        print("illegal argument, try again.")
    elif sys.argv[1] == "--help" or sys.argv[1] == "-h":
        print(""" 
Excel file handler

    A simple program which helps transformed into csv or other file,
     which is much more simple to analysis.Sheets will be saved as name '... - SheetN.csv'
     this will create some files to storage all data
    Due to some problems in Microsoft Excel (R), csv file is encoded as UTF-8 BOM, which
     program cannot write with, you may need to trans file from utf-8 to utf-8 BOM.

    For any issues or bugs, please go to github and report.

    This is an executable version (winPE-32/.exe) which disabled the exit() code.
    For open-source version, please check 
    https://github.com/Holit/XlsxToCsv
-----------------------Copyright (c) 2020 Jerry-------------------------------

usage: eph [-h|-c|-p|--help|--csv|--print] {filename}

    -h  --help  Print this screen
    -c  --csv   trans to .csv (Comma-Separated Values)
                YOU MAY NEED TO TRANS FILE FROM UTF-8 TO UTF-8 BOM.
                THIS IS A BUG OF EXCEL WHEN PROCESSING CSV
    {filename}  file name
        """)
    elif sys.argv[1] == "--csv" or sys.argv[1] == "-c":
        if len(sys.argv) != 3:
            print("illegal input file : null file")
        elif not(os.path.exists(sys.argv[2])):
            print("illegal input file : there's no such file named " + sys.argv[2])
        elif not(os.access(sys.argv[2],os.R_OK)):
            print("illegal input file : access failed at " + sys.argv[2])
        elif not(os.path.splitext(sys.argv[2])[-1][1:] == "xlsx" or os.path.splitext(sys.argv[2])[-1][1:] == "xls"):
            print("illegal input file : is not a excel-based file")
        else:
            try:
                workbook = xlrd.open_workbook(sys.argv[2])
                i = 1
                for table in workbook.sheets():
                    csv_file_name = os.path.splitext(sys.argv[2])[0]
                    csv_file_name = csv_file_name + " - " +  table.name 
                    csv_file_name += ".csv"
                    if(os.path.exists(csv_file_name)):
                        os.remove(csv_file_name)
                    nrows = table.nrows
                    ncols = table.ncols
                    for rows_read in range(1,nrows):
                        row_value = []
                        for cols_read in range(ncols):
                            ctype = table.cell(rows_read, cols_read).ctype
                            nu_str = table.cell(rows_read, cols_read).value
                            if ctype == 2:
                                nu_str = int(nu_str)
                            row_value.append(nu_str)
                        with open(csv_file_name, 'a', encoding='utf-8',newline='') as f:
                            write = csv.writer(f)
                            write.writerow(row_value)
                    i = i +1
                print('Succeed, proceed ' + str(i) + ' sheets \n You may need to use excel to trans encoding from utf-8 to utf-8 BOM')
            except Exception as e:
                print("Unexcepted error occured: " + str(e))