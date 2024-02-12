import openpyxl as op
wbk=op.load_workbook('c:\\openpylearning\\practice\\new_wbk.xlsx')
msheet=wbk.active
for b,c  in msheet['b2:c6']:
       c.value = f'={b.coordinate} * {0.10}'

wbk.save('c:\\openpylearning\\practice\\new_wbk.xlsx')
