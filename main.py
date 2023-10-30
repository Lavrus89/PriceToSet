import xlrd, xlwt
from tkinter.filedialog import askopenfilename
from os.path import basename, dirname, join

# читаем файл и фильтруем
file = askopenfilename()
ws = xlrd.open_workbook(file, encoding_override='1251').sheet_by_index(0)
data = [[ws.cell(row, 0).value, ws.cell(row, 3).value] for row in range(8, ws.nrows-1) if ws.cell(row, 3).value == ' ']

# пишем результат
wb = xlwt.Workbook()
wss = wb.add_sheet('Отчёт')
for row in range(len(data)):
	wss.write(row, 0, data[row][0])
	wss.write(row, 1, data[row][1])
wb.save(join(dirname(file), 'Отчёт_'+basename(file)))
	
	