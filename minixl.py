from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from datetime import datetime
from os.path import basename

doc = "C:\\Users\\patrizio\\Documents\\GitHub\\minixl\\test_data\\t.xlsx"
# doc = "C:\\Users\\Alec\\.projects\\minixl\\test_data\\t.xlsx"
#set the range string which contains the range of header data to be used in hash_year_values
header_cell_range = 'D1:CU1'
wb = load_workbook(filename=doc,use_iterators=True)
sheets = wb.get_sheet_names()



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

def check_pre_event_year():
	'''To find the pre-event year, this function checks
	the 1st event year, minus one, against the values in the
	Data Date column (first 4 digits). The program will iterate up from
	the starting row until it finds a match, or it reaches the end of the
	firms entries. If no match is found, the program will use
	the next event year in the hash_event_years dict, to calculate
	a new pre-event year, and iterate through the Data Dates again. This process repeats
	until there is a match, or the event year list is exausted. If the list is exausted,
	the firm name is appended to a dict (no_pre_event year).
	Once it finds a match, the program checks that same row,column=Net Income(loss)
	for a value. If there is no value, append the firm name to a dict (no_net_income_data) with firm name:pre-event year.
	'''
	ws=wb[sheets[0]]
	event_years = hash_event_years()
	net_income_col = column_index_from_string('AF')-1
	print "Net income col:" + str(net_income_col)
	no_net_income_data = []
	no_first_pre_event_year_data = []
	company_checked = {}
	for row in ws.iter_rows(row_offset=1):
		if row[1].value:
			firm = unicode(row[1].value).strip()
			pre_event_years = [int(x-1) for x in event_years[firm]]
			print firm + " has pre-event years:\n {0}\n".format(pre_event_years)
		if firm not in company_checked and row[13].value:
			datecell = row[13].value
			date = int(str(datecell)[:4])
			net_income = row[net_income_col].value
			if date in pre_event_years:
				if net_income:
					print firm + " has net income data for pre-event year {0}\n\n".format(date)
					company_checked[firm] = date
					if date != min(pre_event_years):
						no_first_pre_event_year_data.append(firm)
				elif firm not in no_net_income_data:
					no_net_income_data.append(firm)
	no_net_income_data = [x for x in no_net_income_data if x not in company_checked]
	no_pre_event_year = [e for e in event_years if e not in company_checked and e not in no_net_income_data]
	total_no_data_entries = no_net_income_data + no_pre_event_year
	print "Firms that have pre-event year data, but no available net income data :\n" + str(no_net_income_data) + "\n\n" + "Firms with no available pre-event year data :\n"+ str(no_pre_event_year)+ "\n"
	with open("log.txt","w") as logfile:
		logfile.write("*" * 50 + "\n" + datetime.now().strftime('%H:%M %d/%m/%Y') + "\n\n" + "Log for analysis of: " + basename(doc) + "\n" + "*" * 50 + "\n\n" + str(len(no_net_income_data)) + " Firms that have pre-event year data, but no available net income data :\n\n")
		for i in sorted(no_net_income_data):
			logfile.write(i + "\n")
		logfile.write("\n" + '*' * 50 + "\n" + str(len(no_pre_event_year)) + " Firms with no available pre-event year data :\n\n")
		for i in sorted(no_pre_event_year):
			logfile.write(i + "\n")
		logfile.write("\n" + "*" * 50 + "\n" + str(len(no_first_pre_event_year_data)) + " Firms for which net income data was not found for the first pre-event year , but was found for a following pre-event year: \n\n")
		for i in sorted(no_first_pre_event_year_data):
			logfile.write(i + "--" + str(company_checked[i]) + "\n")
		print "Results written to log.txt\n\n"
	return total_no_data_entries
	
	
def del_no_data_entries(entries_to_delete):
	''' (list) -> NoneType
	Iterates through the rows of spreadsheet and sets
	all cell values in the row to None if the value in column 2 (firm name)
	matches any value in entries_to_delete.
	Blank cells should be manually removed from the excel sheets
	after running this script. 
	'''
	wb = load_workbook(filename=doc)
	ws=wb[sheets[2]]
	for row in range(2,ws.max_row):
		Cell = ws.cell(column=2,row=row)
		if Cell.value:
			company = unicode(Cell.value).strip()
		if company in entries_to_delete:
			for col in range(2,ws.max_column + 1):
				cell = ws.cell(column=col,row=row)
				cell.value = None				
	wb.save(doc)
	return "All companies with no data have been removed from spreadsheet"
		
		

	
	
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
#write_event_years()
#create_new_xl()
# print check_pre_event_year()
check_pre_event_year()