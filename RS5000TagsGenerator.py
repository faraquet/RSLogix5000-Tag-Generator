# -*- coding: cp1251 -*-
# ѓенерациЯ тегов длЯ RSLogix по теглисту

# !!!!!!!!!!!!!!!!! обращайте внимание на то, какой разделитель строк стоит в региональных настройках !!!!!!!!!!!  

XLS = r'..\taglist.xls' # путь к файлу тег-листа
CSV = r'PLC-Tags_ASKUE.CSV' # путь к csv-файлу

import codecs

f = codecs.open(CSV, 'wt', 'cp1251')

separator = u';' # укажите разделитель строк (такой же, как и в региональных настройках)

template = u'''TAG;;%s;"%s";"%s";"";"(Constant := false, ExternalAccess := Read/Write)"'''
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
for row in range(122, sheet.nrows): #sheet.nrows
    if sheet.cell(row, 3).value == '':
        continue

    tagname = sheet.cell(row, 2).value.strip()
    tagtype = sheet.cell(row, 3).value.strip()
    tagcomment = sheet.cell(row, 1).value.strip()

    if tagtype not in TYPES:
        continue

    comment = repr(tagcomment)[2:-1].replace('\u', '$')
    
    string = template % (tagname, comment, TYPES[tagtype])
    print string
    f.write(string + '\n') 

f.close()
