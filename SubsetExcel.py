from xlrd import open_workbook
from os import path, listdir, chdir
inputLocation ='C:/MergePDF/Input'
dataDrivenFileLoc='C:/MergePDF/Input/DataDrivenFile.xlsx'
chdir(inputLocation)
file_location= dataDrivenFileLoc
# wb = load_workbook(inputLocation + "/" + 'InputOrder.xlsx')
# ws = wb.get_sheet_by_name('InputOrder')
# column_order = ws['A']
# column = ws['B']
workbook= open_workbook(file_location)
worksheet= workbook.sheet_by_index(0)
report_order = worksheet.col(0)
report_name = worksheet.col(1)
type= worksheet.col(2)
list_master = []
for row in range(1, len(report_order)):
    if (str(type[row].value).upper()=='PDF'):
        list_master.append(str(report_name[row].value).strip() + ".pdf")
print(list_master)