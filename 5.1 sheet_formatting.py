import openpyxl as op
from openpyxl.styles import *
from copy import copy
wbk = op.load_workbook('c:\\openpylearning\\practice\\prac1_demo.xlsx')
msheet = wbk['emp_data']
msheet['D20'] = 900
mcell = msheet['D20']


#font formatting
mfont = Font(name='Tahoma', size=18, color=colors.BLUE, bold=True, italic=True, strike=False)
mcell.font = mfont


#pattern Fill
mfill = PatternFill(fill_type='lightGray', fgColor='00FF00')  # '00FF00' is the RGB value for green
mcell.fill = mfill
#border styles
dbl_border_green = Side( border_style='double',color='00FF00')  # '00FF00' is the RGB value for green
thin_border_red = Side(border_style='thin', color='FF0000')  # 'FF0000' is the RGB value for red
mcell.border = Border(left=dbl_border_green, right=thin_border_red, top=dbl_border_green, bottom=thin_border_red)

#Cell Alignment Horzontal and Vertical
align_cell=Alignment(horizontal='left',vertical='bottom')
mcell.alignment=align_cell

#copying formatting styles
# Copying formatting styles
new_mcell = msheet['B2']
new_font = copy(mcell.font)
new_font.color.rgb = '00FF00'  # '00FF00' is the RGB value for green
new_mcell.font = new_font
