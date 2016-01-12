#!python

import xlrd
from xlwt import Workbook

book = Workbook()
sheet1 = book.add_sheet('work')

otchet = xlrd.open_workbook(filename='Отчет.xls', encoding_override='cp1251')
osdv = xlrd.open_workbook(filename='osdv.xlsm')

sheet = otchet.sheet_by_index(0)
dataOtchet = {}
for rownum in range(6, sheet.nrows):
  row = sheet.row_values(rownum)
  mod_row = [row[4], row[5], row[6]]
  dataOtchet[row[2]] = mod_row

dataOsdv = []
sheet = osdv.sheet_by_index(10)
for rownum in range(18, sheet.nrows):
  row = sheet.row_values(rownum)
  if row[5][-1] == '.':
    dataOsdv.append(row[5][:-1])
  else:
    dataOsdv.append(row[5])

ii = 0
for i in dataOsdv:
  ii += 1
#     if i.find(', инв.') != -1:
#         f = i[:i.find(', инв.')]
#         find = dataOtchet.get(f)
#     else:

  if not dataOtchet.get(i):
    find = dataOtchet.get('0' + i)
  else:
    find = dataOtchet.get(i)
  if find:
    sheet1.write(ii, 1, i)
    sheet1.write(ii, 2, find[0])
    sheet1.write(ii, 3, find[1])
    sheet1.write(ii, 4, find[2])

#         if find[0] < 0:
#             x = 0
#             sheet1.write(ii, 1, x)
#             sheet1.write(ii, 2, find[1])
#             sheet1.write(ii, 3, -find[0])
#         else:
#             sheet1.write(ii, 1, find[0])
#             sheet1.write(ii, 2, find[1])
#     else:
#         sheet1.write(ii, 0, i)

book.save('111.xls')
