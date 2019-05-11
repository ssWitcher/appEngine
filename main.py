from flask import Flask, request, send_file
from werkzeug import secure_filename
from PyDictionary import*
import xlrd
import xlwt
import time

app = Flask(__name__)

@app.route('/',methods=['POST'])
def convert_file():
	f = request.files['myfile']
	f.save(secure_filename(f.filename))
	filename = secure_filename(f.filename)
	seconds = time.time()
	wb = xlrd.open_workbook(filename)
	sheet = wb.sheet_by_index(0)
	dictionary=PyDictionary()
	book = xlwt.Workbook(encoding="utf-8")
	sheet1 = book.add_sheet("Sheet 1")
# For row 0 and column 0
	nrows = sheet.nrows
	i = 0
	mean = []
	while i < nrows:
		print(i)
		unformated = sheet.cell_value(i, 0)
		formated = unformated[0:-4]
		mean = dictionary.meaning(formated)
		if mean == None:
			sheet1.write(i,0,unformated)
			sheet1.write(i,1,"No meaning")
		else:
			sheet1.write(i, 0, unformated)
			out = str(mean)
			out = out.replace('{', '')
			out = out.replace('}', '')
			out = out.replace('[', '')
			out = out.replace(']', '')
        #out = out[1:]
			out = out.replace("'",'',2)
			sheet1.write(i,1, out)
		i = i+1
	print(time.time() - seconds)
	book.save("trial11.xls")
	return send_file("trial11.xls")


if __name__ == '__main__':
	app.run()

