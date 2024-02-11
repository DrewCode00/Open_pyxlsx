import openpyxl as op


wbk =op.load_workbook('D:\Data Leads--2024\150k _Fullz.csv')


#msheet =wbk['emp_data']

#msheet['D20']

#new_item=('duncan', 25, 'ipads', 3000)
#msheet.append(new_item)

#wbk.save('D:\Data Leads--2024\150k _Fullz.csv')


mcell=msheet['e2']="bonus"

for salary,bonus in msheet['d2:e9']:
    bonus.value= salary.value * 0.25
