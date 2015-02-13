# -*- coding:utf-8 -*-
__author__ = 'fantasy'

import xlrd
import os, sys, glob


ROW_START = 8
COL_START = 1
#COL_END = 10


FIRST_LINE = 'For M5|Adr{0:>6}|TO  {1:<27}|{2:22}|{3:22}|{4:22}|'
SECOND_LINE = 'For M5|Adr{0:>6}|TO  Start-Line{1:7}{2:<9}1|{3:22}|{4:22}|{5:22}|'
THIRD_LINE = 'For M5|Adr{0:>6}|KD1{1:>9}{2:18}1|{3:22}|{4:22}|Z{5:>16} m   |'
ITEM_RB_LINE = 'For M5|Adr{0:>6}|KD1{1:>9}{2:6}{3:<12}1|Rb{4:>15} m   |HD{5:>15} m   |{6:22}|'
ITEM_RF_LINE = 'For M5|Adr{0:>6}|KD1{1:>9}{2:6}{3:<12}1|Rf{4:>15} m   |HD{5:>15} m   |{6:22}|'
ITEM_FOOTER_LINE = 'For M5|Adr{0:>6}|KD1{1:>9}{2:6}{3:<12}1|{4:22}|{5:22}|Z{6:>16} m   |'

PAGE_FOOTER_FIRST_LINE = 'For M5|Adr{0:>6}|KD1{1:>9}{2:18}1|{3:<2}{4:15} m   |{5:<2}{6:>15} m   |Z{7:>16} m   |'
PAGE_FOOTER_SECOND_LINE = 'For M5|Adr{0:>6}|KD1{1:>9}{2:18}1|{3:<2}{4:15} m   |{5:<2}{6:>15} m   |Z{7:>16} m   |'
PAGE_FOOTER_THIRD_LINE = 'For M5|Adr{0:>6}|TO  End-Line{1:18}1|{2:22}|{3:22}|{4:22}|'
'''
For M5|Adr     1|TO  1215bbb.dat                |                      |                      |                      |
For M5|Adr     2|TO  Start-Line       BBFF     1|                      |                      |                      |
For M5|Adr     3|KD1      BM1                  1|                      |                      |Z         0.00000 m   |
----------------------------------------------------------------------------------------------------------------------
For M5|Adr     4|KD1      BM1      01:33:402   1|Rb        1.85953 m   |HD         45.870 m   |                      |
For M5|Adr     5|KD1      BM1      01:34:472   1|Rb        1.85957 m   |HD         46.426 m   |                      |
For M5|Adr     6|KD1       TP      01:35:392   1|Rf        1.46981 m   |HD         45.290 m   |                      |
For M5|Adr     7|KD1       TP      01:35:462   1|Rf        1.46987 m   |HD         45.315 m   |                      |
For M5|Adr     8|KD1       TP      01:35:46    1|                      |                      |Z         0.38971 m   |
----------------------------------------------------------------------------------------------------------------------
For M5|Adr    34|KD1      BM1                  1|Sh        0.00019 m   |dz       -0.00019 m   |Z         0.00000 m   |
For M5|Adr    35|KD2      BM1        6         1|Db         240.26 m   |Df         236.36 m   |Z         0.00019 m   |
For M5|Adr    36|TO  End-Line                  1|                      |                      |                      |
'''


def getExcleFiles(path):
    files = list()
    if os.path.isfile(path):
        return files.append(path)
    elif os.path.isdir(path):
        for file_path in glob.glob(path + os.sep + '*.xls'):
            files.append(file_path)
        return files
    elif os.path.exists(path):
        print 'The Path is not exist,please check it!'
        return files
    else:
        print 'Unknown Errors，please contact fantasy!'
        return files


def getDataFromExcle(file_path):
    data = list()
    footer = list()
    try:
        rb = xlrd.open_workbook(file_path)
        rs = rb.sheet_by_index(0)

        nrows = rs.nrows
        ncols = rs.ncols

        item = list()

        for row in range(ROW_START, nrows - 5):
            for col in range(COL_START, ncols):
                if col == 1 and not rs.cell_value(row, col).strip():
                    #print row,col,':', rs.cell_value(row,col)
                    #row += 1
                    break
                    #continue
                else:
                    value = rs.cell_value(row, col)
                    if str(value).strip():
                        item.append(value)

            if len(item):
                data.append(item)
                item = list()

        for row in range(nrows - 3, nrows):
            item = list()
            for col in range(COL_START, ncols - 4):
                item.append(rs.cell_value(row, col))

            footer.append(item)
    except Exception, msg:
        print msg

    return data, footer


def formatData(data, footer):
    file_name = os.path.basename(file_path)
    file_ext = file_name.split('.')[-1]
    dat_file = file_name.replace(file_ext, 'DAT')

    first_line_format = ('1', dat_file, '', '', '')
    second_line_format = ('2', '', 'BBFF', '', '', '')
    third_line_format = ('3', data[0][0], '', '', '', '0.00000')

    line_index = 4

    result_list = list()

    result_list.append(FIRST_LINE.format(*first_line_format))
    result_list.append(SECOND_LINE.format(*second_line_format))
    result_list.append(THIRD_LINE.format(*third_line_format))



    for d in data:
        if data.index(d) == 0:
            item_line_format = (line_index, d[0], '', d[6], d[3], d[1], '')
            result_list.append(ITEM_RB_LINE.format(*item_line_format))
            line_index += 1

            item_line_format = (line_index, d[0], '', d[7], d[4], d[2], '')
            result_list.append(ITEM_RB_LINE.format(*item_line_format))
            line_index += 1

        elif data.index(d) == 1:
            item_line_format = (line_index, d[0], '', d[6], d[3], d[1], '')
            result_list.append(ITEM_RF_LINE.format(*item_line_format))
            line_index += 1

            item_line_format = (line_index, d[0], '', d[7], d[4], d[2], '')
            result_list.append(ITEM_RF_LINE.format(*item_line_format))
            line_index += 1

            item_footer_format = (line_index, d[0], '', d[7][0:-1], '', '', d[5])
            result_list.append(ITEM_FOOTER_LINE.format(*item_footer_format))
            line_index += 1

        else:
            if len(d) == 7:
                item_line_format = (line_index, d[0], '', d[5], d[3], d[1], '')
                result_list.append(ITEM_RB_LINE.format(*item_line_format))
                line_index += 1

                item_line_format = (line_index, d[0], '', d[6], d[4], d[2], '')
                result_list.append(ITEM_RB_LINE.format(*item_line_format))
                line_index += 1

            else:
                item_line_format = (line_index, d[0], '', d[6], d[3], d[1], '')
                result_list.append(ITEM_RF_LINE.format(*item_line_format))
                line_index += 1

                item_line_format = (line_index, d[0], '', d[7], d[4], d[2], '')
                result_list.append(ITEM_RF_LINE.format(*item_line_format))
                line_index += 1

                item_footer_format = (line_index, d[0], '', d[7][0:-1], '', '', d[5])
                result_list.append(ITEM_FOOTER_LINE.format(*item_footer_format))
                line_index += 1



    page_footer_first_format = (line_index, footer[0][1], '', footer[1][0], footer[1][1], footer[1][2], footer[1][3],
                                str(footer[1][5]).ljust(7, '0'))
    page_footer_second_format = (line_index, footer[0][1], '', footer[2][0], footer[2][1], footer[2][2], footer[2][3],
                                 str(footer[2][5]).ljust(7, '0'))
    page_footer_third_format = (line_index, '', '', '', '')

    result_list.append(PAGE_FOOTER_FIRST_LINE.format(*page_footer_first_format))
    line_index += 1
    result_list.append(PAGE_FOOTER_SECOND_LINE.format(*page_footer_second_format))
    line_index += 1
    result_list.append(PAGE_FOOTER_THIRD_LINE.format(*page_footer_third_format))

    return result_list


if __name__ == '__main__':
    input_path = ''
    output_path =''
    config_path = os.path.join(os.path.abspath(os.path.dirname(sys.argv[0])), 'config.ini')
    config = open(config_path,'r')
    try:
        input_path = config.readline().split('=')[-1].strip()
        output_path = config.readline().split('=')[-1].strip()
        #print input_path,output_path
    finally:
        config.close()

    #path = '/Users/fantasy/Desktop/'
    files = getExcleFiles(input_path)

    for file_path in files:
        data, footer = getDataFromExcle(file_path)
        result_list = formatData(data,footer)

        file_name = os.path.basename(file_path)
        file_ext = file_name.split('.')[-1]
        dat_file = file_name.replace(file_ext, 'DAT')

        with open(output_path + os.path.sep + dat_file,'w') as f:
            f.writelines([line+'\n' for line in result_list])

        print file_path +' transform complete!'

        print '******************************************************'