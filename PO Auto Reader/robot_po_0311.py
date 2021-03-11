''' start '''

import sys
import os
import xlwt,xlrd
import re
from xlutils.copy import copy
from datetime import datetime
import locale


FILE_NAME = []
XL_NAME = []
def find_files(files):
    for file in files:
        if '.htm' in file:
            file = file.replace("\'", "")
            FILE_NAME.append(file)
        if file[-4:] == '.xls':
            print(file)
            file = file.replace("\'", "")
            XL_NAME.append(file)

allfiles = os.listdir('.')
#print(allfiles)
find_files(allfiles)
#print(FILE_NAME, XL_NAME[0])
locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
#print(locale.getlocale(),datetime.today().strftime('%d %b. %Y'))
def read_xl(list_po, i, length):
    data_excel=xlrd.open_workbook(XL_NAME[0], formatting_info=True)
    table=data_excel.sheets()[0]
    n_rows=table.nrows  # 获取该sheet中的有效行数
    n_cols=table.ncols  # 获取该sheet中的有效列数
    #print(n_cols,n_rows)
    write_xl(n_rows, list_po, data_excel, i, length)

def write_xl(n_rows, list_po, data_excel, index, length):
    workbook = copy(data_excel)
    pos = 0
    for value in list_po:
        if pos == 10 and index == length - 1:
            workbook.get_sheet(0).write_merge(n_rows - length + 1, n_rows, pos, pos, value)
        elif pos == 10 and index != length:
            pass
        else:
            workbook.get_sheet(0).write(n_rows, pos, value)  # 不带样式的写入
        pos += 1
    workbook.save(XL_NAME[0])

def clean(liste):
    while '' in liste:
        liste.remove('')
    while '&nbsp;' in liste:
        liste.remove('&nbsp;')
    return liste

def rd_file(file_name):
    po_no = file_name.split(".")[0]
    list_po = []
    #print(po_no)
    fopen = open(file_name, "r", encoding="utf-8")
    lines = fopen.readlines()
    i = 0
    shop_name = ''
    po_date = ''
    model_livdate = []
    quantity = []
    index = []
    addr = []
    total = []
    dest = []
    for i in range(len(lines)):
        if re.findall("fdml-ov-bill-to-gf-addr-name-val", lines[i]) and shop_name == '':
            shop_name = lines[i+1].split('\n')[0]
            #print(shop_name)
        if re.findall("Commande soumise le :", lines[i]) and po_date == '':
            po_date = ' '.join(lines[i].split(' ')[4:7])
            po_date = datetime.strptime(po_date, '%d %b %Y')
            #po_date = po_date.strftime('%Y/%m/%d')
            po_date = po_date.strftime('%d/%m/%Y')
            today = datetime.today().strftime('%Y/%m/%d')
            #print(today)
        if re.findall('<td style="vertical-align: top; vertical-align: top; " valign=top align=left colspan=1 class="tableBodyClass fdml-ov-itemOutTable vr-forceVerticalAlignTop">', lines[i]):
            model_livdate.append(lines[i + 1].replace('\n', ''))
            model_livdate.append(lines[i + 2].replace('\n', ''))
            model_livdate.append(lines[i + 3].replace('\n', ''))
        if re.findall('<td colspan=1 align=left valign=top style="vertical-align: top; " class="tableBodyClass fdml-ov-itemOutTable vr-forceVerticalAlignTop">', lines[i]):
            quantity.append(lines[i + 10].replace('\n', ''))
        if re.findall('<td style="vertical-align: top; vertical-align: top; " colspan=1 valign=top align=left class="cus-ncdv-formpadvalue ANXLabel">', lines[i]):
            dest.append(lines[i + 3])
        if re.findall('<span style="vertical-align: top; font-weight:bold">', lines[i]):
            addr.append(lines[i + 1])
        if re.findall('<span style="vertical-align: top; ">', lines[i]):
            index.append(i)
        if re.findall('<td style="vertical-align: top; " valign=top align=right colspan=1 class="tableBodyClass fdml-ov-itemOutTable vr-forceVerticalAlignTop">', lines[i]):
            total.append(lines[i + 3].replace(' EUR\n', '').replace(u'\xa0', '').replace(',', '.'))
            
    #print("total", total)
    addr.append(lines[int(index[1])+3].replace('<br>', ''))
    addr.append(lines[int(index[1])+7].replace('\n', ''))
    addr.append(lines[int(index[1])+8])
    addr.append(lines[int(index[1])+13])
    addr.append(lines[int(index[1])+40].split('&')[0])
    addr.append(lines[int(index[1])+45])
    addr.append(lines[int(index[1])+138].split('&')[0])
    addr.append(lines[int(index[1])+142].split('>')[1])
    addr.insert(0, dest[0])
    addr = ''.join(addr)
    print(addr)
    model_livdate = [x for x in model_livdate if x != '\n']
    model_livdate = [x for x in model_livdate if x != '&nbsp;\n']
    clean(model_livdate)
    products = model_livdate
    #print(products)
    i = 0
    for index in range(0, int(len(products)/2)):
        product = products[::2]
        dateliv = products[1::2]
        # print(product, dateliv)
        if '@' in product[i]:
            name_p = 'ESL'
            quantity_real = int(quantity[i]) * 200
            price = float(total[i]) / quantity_real
        elif 'Frais' in product[i]:
            name_p = 'DELIVERY'
            quantity_real = int(quantity[i])
            price = total[i]
        elif 'Access' in product[i]:
            name_p = 'AP'
            quantity_real = int(quantity[i])
            price = total[i]
        elif 'ShopWeb' in product[i]:
            name_p = 'SOFTWARE'
            quantity_real = int(quantity[i])
            price = total[i]
        else:
            name_p = '其他费用'
            quantity_real = int(quantity[i])
            price = total[i]
        #print(price, quantity_real)
        #list_po = [today, po_no, 'France', 'AUCHAN', shop_name, po_date, name_p, product[i], quantity[i], quantity_real, dateliv[i], '', '', '', '', '', addr, 'EUR', str(price), str(total[i])]
        price = str(price).replace('.', ',')
        list_po= [shop_name, po_no, 'AUCHAN', 'HANSHI',  po_date, name_p, product[i], quantity_real, str(price), dateliv[i], addr, str(total[i])]
        read_xl(list_po, i, int(len(products)/2))
        i += 1
    #print(int(len(products)/2))

for file in FILE_NAME:
    rd_file(file)
exit()