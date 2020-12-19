import xlrd
import csv
import sys
import os

csv_file_name = 'excel_to_csv.csv'

def read_excel(excel_name):
  #读取Excel文件每一行内容到一个列表中
  workbook = xlrd.open_workbook(excel_name)
  table = workbook.sheet_by_index(0) #读取第一个sheet
  nrows = table.nrows
  ncols = table.ncols
  # 跳过表头，从第一行数据开始读
  for rows_read in range(1,nrows):
    #每行的所有单元格内容组成一个列表
    row_value = []
    for cols_read in range(ncols):
      #获取单元格数据类型
      ctype = table.cell(rows_read, cols_read).ctype
      #获取单元格数据
      nu_str = table.cell(rows_read, cols_read).value
      #判断返回类型
      # 0 empty,1 string, 2 number(都是浮点), 3 date, 4 boolean, 5 error
      #是2（浮点数）的要改为int
      if ctype == 2:
        nu_str = int(nu_str)
      row_value.append(nu_str)
    yield row_value
 
def xlsx_to_csv(csv_file_name,row_value):
  #生成csv文件
  with open(csv_file_name, 'a', encoding='utf-8',newline='') as f: #newline=''不加会多空行
    write = csv.writer(f)
    write.writerow(row_value)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        print("illegal argument, try again.")
        #exit(-1)
    elif len(sys.argv[1]) == 0:
        print("illegal argument, try again.")
        #exit(-1)
    elif sys.argv[1] == "--help" or sys.argv[1] == "-h":
        print(""" 
Excel file handler

    A simple executable program which helps transformed into csv or other file,
    which is much more simple to analysis.

    This is an executable version which disabled the exit() code.
    For opensource version, please check github.
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
            #exit(-2)
        elif not(os.path.exists(sys.argv[2])):
            print("illegal input file : there's no such file named " + sys.argv[2])
            #exit(-3)
        #This is not strict for win32 platform, which may cause errors.
        elif not(os.access(sys.argv[2],os.R_OK)):
            print("illegal input file : access failed at " + sys.argv[2])
            #exit(-4)
        elif not(os.path.splitext(sys.argv[2])[-1][1:] == "xlsx" or os.path.splitext(sys.argv[2])[-1][1:] == "xls"):
            print("illegal input file : is not a excel-based file")
            #exit(-5)
        else:
            try:
                #creating csv file path
                csv_file_name = os.path.splitext(sys.argv[2])[0] + ".csv"
                for row_value in read_excel(sys.argv[2]):
                  xlsx_to_csv(csv_file_name,row_value)
                print('Succeed \n You may need to use excel to trans encoding from utf-8 to utf-8 BOM')
                #exit(0)
            except Exception as e:
                print("Unexcepted error occured: " + str(e))
                #exit(-6)