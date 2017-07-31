#! /usr/bin/env python3

import os
import xlwt
import readtxt


class trans2excel():
    '''
    trans text files to a excel file(.xls)
    __init__:
    *args：路径
    **kwargs:
    output:输出prefix，默认：./tmp
    sep:分隔符，默认：\t
    type:输出类型，默认：xls（暂时只能输出这种）

    '''
    def __init__(self, *path, output='./tmp', sep='\t', type='xls'):
        self._path = path
        self._out = output
        self._sep = sep
        self._type = type
        self._files = []

    def _isfile(self, mydir):
        for lists in os.listdir(mydir):
            path = os.path.join(mydir, lists)

            if os.path.isdir(path):
                self._isfile(path)
            else:
                self._files.append(path)

    def _getfiles(self):
        for path in self._path:
            if os.path.isdir(path):
                self._isfile(path)
            else:
                self._files.append(path)
            return self._files

    def _extractlines(self, file):
        rt = readtxt.readtxt(file, self._sep)
        lines = rt.readfile()
        return lines

    def generateExcel(self):
        self._getfiles()
        if not os.path.exists('{}.{}'.format(self._out, self._type)):
            workbook = xlwt.Workbook(encoding='ascii')
            for efile in self._files:
                filename = os.path.basename(efile)
                filename = filename.replace('.txt', '')
                worksheet = workbook.add_sheet(filename)
                mylines = self._extractlines(efile)
                row = 0
                for line in mylines:

                    line = line.strip()
                    cells = line.split(self._sep)
                    for col in range(0, len(cells)):
                        worksheet.write(row, col, label=cells[col])
                    row += 1

            workbook.save('{}.{}'.format(self._out, self._type))
        else:
            print('({}.{}) is existed, please check the File!!\n'
                  .format(self._out, self._type))


if __name__ == '__main__':
    pass
    # test = trans2excel('/annoroad/data1/bioinfo/PROJECT/Commercial/Medical/'
    #                    'Leukemia/data/Commercial/HB_234_20170318105823/result/'
    #                    'HB15CK00215-1-32/Variant/SNP-INDEL_MT/FILTER')
    # files = test.getfiles()
    # test.generateExcel()
    # for i in files:
    #     print('file:{}'.format(i))
