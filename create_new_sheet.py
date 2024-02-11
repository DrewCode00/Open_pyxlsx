import openpyxl as op

wbk =op.workbook()
msheet = wbk.active()
profit = { 'yr 2019':50000, 'Yr 2020':75000, 'yr 2023':97000, 'yr 2023':19200}
msheet['A1']='Year'
msheet['B1']= 'profit'
for k.p in profit.items():
    msheet.append(k.p)

    wbk.save('D:/Data Leads--2024/new_wbk.xlsx')