import os
from openpyxl import load_workbook
from openpyxl import Workbook

doc = "C:\\Users\\Alec\\.projects\\minixl\\test_data\\t.xlsx"
#set the range string which contains the range of header data to be used in hash_year_values
header_cell_range = 'D1:CU1'
wb = load_workbook(filename=doc,use_iterators=True)
sheets = wb.get_sheet_names()
print sheets

def get_company_names():
	ws = wb[sheets[0]]
	result = []
	for row in ws.iter_rows(row_offset=1):
		if row[1].value:
			company = row[1].value.strip()
			result.append(company)
	return result



def hash_year_values():
	ws=wb[sheets[1]]
	result = {}
	for row in ws.iter_rows(range_string = header_cell_range):
		for cell in row:
			value = cell.value
			if isinstance(value,long) or isinstance(value,int):
				result[cell.column] = value
	return result

def hash_event_years():
	ws=wb[sheets[1]]
	result={}
	company_list = get_company_names()
	year_dict = hash_year_values()
	for row in ws.iter_rows(row_offset=1):
		company_name = row[0].value.strip()
		event_years =[]
		if company_name in company_list:
			for cell in row:
				if cell.column in year_dict and cell.value:
					event_years.append(year_dict[cell.column])
			event_years.sort()
			result[company_name] = event_years
	return result


def write_event_years():
	wb = load_workbook(filename=doc)
	ws=wb[sheets[0]]
	event_year_dict = hash_event_years()
	companies = []
	prevcompany = ''
	for row in range(2,ws.max_row):
		Cell = ws.cell(column=2,row=row)
		company = unicode(Cell.value).strip()
		if company and company != prevcompany and company in event_year_dict:
			ws.cell(column=3,row=row).value = event_year_dict[company][0]
			companies.append(company)
			prevcompany = company
	wb.save(doc)
	print "New values written to spreadsheet"
	return companies

# def check_pre_event_year():


# def create_new_xl():
	# nb = Workbook(write_only=True)
	# ws = nb.create_sheet()
	# entry_list = hash_event_years()
	# for entry in entry_list:
		# ws.append([entry,entry_list[entry],entry_list[entry]-1])
	# nb.save('C:\\Users\\patrizio\\Documents\\GitHub\\minixl\\test_data\\event_years.xlsx')
	# print "Saved new workbook with eventyears"

# companies_sheet1 = get_company_names()
# companies_sheet2 = hash_event_years()
# for i in companies_sheet1:
	# if i not in companies_sheet2:
		# print i
# print len(get_company_names())
# print hash_event_years()
# print hash_year_values()
write_event_years()
#create_new_xl()