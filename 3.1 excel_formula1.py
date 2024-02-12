import openpyxl as op
wbk=op.load_workbook('c:\\openpylearning\\practice\\new_wbk.xlsx')
msheet=wbk.active
msheet['c1']="Commission"
msheet['c2']=  '=b2*0.10'
wbk.save('c:\\openpylearning\\practice\\new_wbk.xlsx')

