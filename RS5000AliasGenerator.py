# -*- coding: cp1251 -*-
# ѓенерациЯ алиасов длЯ RSLogix по теглисту

# !!!!!!!!!!!!!!!!! обращайте внимание на то, какой разделитель строк стоит в региональных настройках !!!!!!!!!!!  

XLS = r'..\taglist.xls' # путь к файлу тег-листа
CSV = r'PLC-Tagsakue.CSV' # путь к csv-файлу

import codecs

f = codecs.open(CSV, 'wt', 'cp1251')

separator = u';' # укажите разделитель строк (такой же, как и в региональных настройках)

template = u'''ALIAS;;%s;"%s";"";"%s";"(Constant := false, ExternalAccess := Read/Write)"''' # для алиасов
template = template.replace(';', separator)

header = '''remark;"CSV-Import-Export"
0.3
TYPE;SCOPE;NAME;DESCRIPTION;DATATYPE;SPECIFIER;ATTRIBUTES'''
header = header.replace(';',  separator)

f.write(header + '\n')

import xlrd

rbook = xlrd.open_workbook(XLS,  encoding_override='cp1251')
sheet_name = u'TAGS'
TYPES = {'AI':'_AI','DI':'_DI','DO':'_DO','AO':'PID','VALVE':'_VALVE','ENGINE':'_ENGINE'}

if not sheet_name in rbook.sheet_names():
    print 'warning! file have not sheet',  sheet_name                
print 'Analysing sheet',  sheet_name, '...\n',
sheet = rbook.sheet_by_name(sheet_name)

print header
for row in range(120, 293): #sheet.nrows
    if sheet.cell(row, 3).value == '':
        continue

    tagname = sheet.cell(row, 2).value.strip()
    tagtype = sheet.cell(row, 3).value.strip()
    shassi  = sheet.cell(row, 4).value
    slot    = sheet.cell(row, 5).value
    point   = sheet.cell(row, 6).value
    if tagtype == u'AI':        
        alias = shassi + ':' + str(int(slot)) + ':I.Ch' + str(int(point)) + 'Data'
    if tagtype == u'DI':
        alias = shassi + ':' + str(int(slot)) + ':I.Data.' + str(int(point))
    if tagtype == u'DO':
        alias = shassi + ':' + str(int(slot)) + ':I.Data.' + str(int(point))    
    tagcomment = sheet.cell(row, 1).value.strip()

    if tagtype not in TYPES:
        continue

    comment = repr(tagcomment)[2:-1].replace('\u', '$')
    
    string = template % (tagname, comment, alias)
    print string
    f.write(string + '\n') 

f.close()
