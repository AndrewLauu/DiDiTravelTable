# -*- coding: utf-8 -*-
# from bs4 import BeautifulSoup as bs
import openpyxl as xl
from itertools import islice
from nodeTmp import etNode
import xml.etree.ElementTree as et

# from nodeTmp import bsNode

print('Loaded libs.')

# 读取模板
workbook = xl.load_workbook('files\\template.xlsx', read_only=True, data_only=True)
worksheet = workbook['tmp']
# worksheet = workbook.worksheets[0]
nRow = worksheet.max_row
workTable = worksheet.iter_rows(min_row=2, values_only=True)
# priceCol = worksheet.iter_rows(min_col=8, max_col=8, min_row=2, values_only=True)
# priceCol = list(zip(*priceCol))[0]
# sumPrice = sum([float(i) for i in priceCol])
sumPrice = workbook['价格']['a13'].value

# 判断行数，拆分页码（分表）
# pg0:0 -> 13r
# pg1:13 + 18 * (1-1) -> 13 + 18 * 1
# pg2:13+18 * (2-1) -> 13 + 18 * 2
nPage = (nRow - 1 - 13) // 18 + 1
table0 = islice(workTable, 0, 13)
tables = {p: islice(workTable, 13 + 18 * p - 18, 13 + 18 * p)
          for p in range(1, nPage)}
tables.update({0: table0})
tables = dict(sorted(tables.items(), key=lambda d: d[0]))

# 注册命名空间，实例化doc
# doc = bs(docTmp, 'xml')
for prefix, uri in etNode.xmlns:
    et.register_namespace(prefix, uri)
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
doc = et.fromstring(etNode.docTmp)

# 增加信息段
# @nRow
# @sumPrice
pInfoTmp = etNode.pInfoTmp.replace('@nRow', f'{nRow - 1}').replace('@sumPrice', str(sumPrice))
# pInfo = bs(pInfoTmp, 'xml')
pInfo = et.fromstring(pInfoTmp)
doc.find('w:body').append(pInfo)
print('Read excel.')

# 实例化表格及行元素
# tbl = bs(tblTmp, 'xml')
# tr = bs(trTmp, 'xml')
tbl = et.fromstring(etNode.tblTmp)
tr = et.fromstring(etNode.trTmp)

# 实例化padding, blankRow
# pPadding = bs(pPaddingTmp, 'xml')
# pBlank = bs(pBlankTmp, 'xml')
pPadding = et.fromstring(etNode.pPaddingTmp)
pBlank = et.fromstring(etNode.pBlankTmp)

# 建立偶行格式
# w_shd_0 = tr.new_tag('w:shd', attrs={'w:val' :"clear", 'w:color':"auto",'w:fill':"666666"})
# w_shd_even = tr.new_tag('w:shd', attrs={'w:val':  'clear', 'w:color': 'auto',
#                                         'w:fill': 'F0F0F0'})
w_shd_even = et.SubElement(tr, 'w:shd', attrib={'w:val': 'clear', 'w:color': 'auto', 'w:fill':
                                                         'F0F0F0'})

# 遍历表中的分表，完善表元素并插入doc
for page, table in tables.items():
    print(f'scanning table {page}')
    # 遍历分表中的行，完善行元素并插入tbl
    for row in table:
        # 实例化单元格元素，判断奇偶行,匹配格式
        # tc = bs(tcTmp, 'xml')
        tc = et.fromstring(etNode.tcTmp)
        print(f'Scanning row {row[0]}')
        if row[0] % 2 == 0:
            tc.find(w + 'tcPr').append(w_shd_even)

        # 遍历行中单元格，完善单元格元素，并插入tr
        for cell in row:
            # w_t = tc.new_tag('w:t')
            # w_t.string = str(cell)
            w_t = et.SubElement(tc, 'w:t')
            w_t.text = str(cell)
            tc.find(w + 'r').append(w_t)
            tr.append(tc)
        # tr插入tbl
        tbl.append(tr)
    # 完善tbl，插入doc
    # if page==0:
    print('Adding prefix')
    doc.find('body').append(pPadding)
    doc.find('body').append(pBlank)
    doc.find('body').append(tbl)

# 实例化sectPr并插入doc
# sectPr = bs(sectPrTmp, 'xml')
sectPr = et.fromstring(etNode.sectPrTmp)
doc.find('body').append(sectPr)
workbook.close()
print('write to file')
with open('files/DiDiTravelPersonnel/word/document.xml', 'w') as f:
    f.write(str(doc))
