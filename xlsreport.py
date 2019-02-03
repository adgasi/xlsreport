import sys, os, getopt
from time import time
from datetime import datetime

from xlutils.copy import copy
from xlrd import open_workbook
import xlwt

def main():
    # parse and validate CL arguments
    argv = sys.argv[1:]
    srcPath = ''
    try:
        opts, args = getopt.getopt(argv, "i:")
        if not opts:
            usage()
            sys.exit(2)
    except getopt.GetoptError:
        usage()
        sys.exit(2)

    for opt, arg in opts:
        print('for loop!!', opt, arg)
        if opt == '-h':
            usage()
            sys.exit(2)
        elif opt == '-i':
            srcPath = arg
    print('srcPath', srcPath)
    # xlwt creates columns with a default width of 2962. 
    # measured in 1/256 of the width of the character ‘0’. 
    COL_BASE_WIDTH = 256

    book = copy(open_workbook(srcPath))
    sheet = book.get_sheet(0)

    # styles
    styleCommon = xlwt.XFStyle()
    styleHeaders = xlwt.XFStyle()
    
    styleCommon.borders = get_borders(1)
    styleHeaders.borders = get_borders(1)

    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    styleHeaders.alignment = alignment

    # write to cells
    sheet.write(9, 12, '　　年　　月　　日')
    sheet.write(10, 10, '本人', styleHeaders)
    sheet.write(10, 11, '部長', styleHeaders)
    sheet.write(10, 12, '日付', styleHeaders)

    # merged cells
    sheet.write_merge(11, 12, 10, 10, '', styleCommon)
    sheet.write_merge(11, 12, 11, 11, '', styleCommon)
    sheet.write_merge(11, 12, 12, 12, '', styleCommon)

    # resize
    signCol = sheet.col(10)
    signCol.width = COL_BASE_WIDTH * 20
    dateCol = sheet.col(12)
    dateCol.width = COL_BASE_WIDTH * 20

    dataRow = sheet.row(12)
    dataRow.height_mismatch = True
    dataRow.height = COL_BASE_WIDTH * 2

    # create dir and save xls
    ts = time()
    dirname = datetime.fromtimestamp(ts).strftime('%Y%m%d%H%M%S')
    filename = '[working]monthly_individual_working_list.xls'
    dirpath = './output/'+dirname

    os.makedirs(dirpath)
    book.save(dirpath+'/'+filename)

def usage():
    print('Usage: xlsreport.py -i <inputfile>')

def get_borders(size):
    borders = xlwt.Borders()
    borders.left = 1
    borders.top = 1
    borders.right = 1
    borders.bottom = 1

    return borders

if __name__ == '__main__':
    main()