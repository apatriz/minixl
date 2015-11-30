from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string
from datetime import datetime
from os.path import basename
import itertools

old_doc = "C:\\Users\\Alec\\.projects\\minixl\\test_data\\t.xlsx"
doc = "C:\\Users\\Alec\\.projects\\minixl\\test_data\\3_Target_firm.xlsx"
industry_firms = "C:\\Users\\Alec\\.projects\\minixl\\test_data\\2_BigData.xlsx"
#set the range string which contains the range of header data to be used in hash_year_values
header_cell_range = 'D1:CU1'
wb = load_workbook(filename=doc,use_iterators=True)
sheets = wb.get_sheet_names()
wb2 = load_workbook(filename=industry_firms,use_iterators=True)
wb3 = load_workbook(filename=old_doc,use_iterators=True)
sheets_old = wb3.get_sheet_names()




def get_company_names():
	ws = wb[sheets[2]]
	result = []
	for row in ws.iter_rows(row_offset=1):
		if row[1].value:
			company = row[1].value.strip()
			result.append(company)
	return result



def hash_year_values():
	ws=wb3[sheets_old[1]]
	result = {}
	for row in ws.iter_rows(range_string = header_cell_range):
		for cell in row:
			value = cell.value
			if isinstance(value,long) or isinstance(value,int):
				result[cell.column] = value
	return result

def hash_event_years():
	ws=wb3[sheets_old[1]]
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
	ws=wb[sheets[2]]
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
	changed_event_year = {i:company_checked[i] for i in no_first_pre_event_year_data}
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
	return changed_event_year


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

def change_entry(entries_to_change):
	''' (list) -> NoneType

	'''
	wb = load_workbook(filename=doc)
	ws=wb[sheets[2]]
	for row in range(2,ws.max_row):
		Cell = ws.cell(column=2,row=row)
		if Cell.value:
			company = unicode(Cell.value).strip()
			pre_event_year = ws.cell(column=4,row=row)
			event_year = ws.cell(column=3,row=row)
			if company in entries_to_change:
				pre_event_year.value = int(entries_to_change[company])
			else:
				pre_event_year.value = event_year.value - 1
	wb.save(doc)
	return "Changed event years"


def build_target_firm_data():
	sic_col = column_index_from_string('BW') -1
	event_col = column_index_from_string('C') - 1
	pre_event_col = column_index_from_string('D') - 1
	company_col = column_index_from_string('B') - 1
	assets_col = column_index_from_string('Y') - 1
	net_income_col = column_index_from_string('AF') - 1
	date_col = column_index_from_string('N') - 1

	ws = wb[sheets[2]]
	result = {}
	for row in ws.iter_rows(row_offset=1):
		datecell = row[date_col].value
		if datecell:
			date = int(str(datecell)[:4])
		if row[company_col].value:
			company = unicode(row[company_col].value).strip()
			sic_code = row[sic_col].value
			eventyear = row[event_col].value
			pre_eventyear = row[pre_event_col].value
			result[company] = {"sic_code":sic_code,"eventyear":eventyear,"pre_eventyear":pre_eventyear}
		if date == eventyear:
			total_assets = row[assets_col].value
			if total_assets:
				result[company]["total_assets"] = total_assets
			else:
				result[company]["total_assets"] = 0
		if date == pre_eventyear:
			net_income = row[net_income_col].value
			if net_income:
				result[company]["net_income"] = net_income
			else:
				result[company]["net_income"] = 0
	return result

def build_industry_groups():
	ws = wb2.active
	target_firms = build_target_firm_data()
	result = {}
	company_col = column_index_from_string('I') - 1
	sic_col = column_index_from_string('BK') - 1
	date_col = column_index_from_string('B') - 1
	assets_col = column_index_from_string('M') - 1
	net_income_col = column_index_from_string('T') - 1
	for row in ws.iter_rows(row_offset=1):
		datecell = row[date_col].value
		date = int(str(datecell)[:4])
		sic_code = row[sic_col].value
		firm = unicode(row[company_col].value).strip()

		firm_total_assets = row[assets_col].value
		net_income = row[net_income_col].value
		if not firm_total_assets or not net_income:
			continue
		#TODO: need to fix name matching for edge cases
		target_firm_names = set([name for name in target_firms if target_firms[name]["sic_code"] == sic_code and name.upper()[:-1] not in firm and firm[:-1] not in name.upper()])
		for name in target_firm_names:
			if date == target_firms[name]["eventyear"] and "total_assets" in target_firms[name] and 0.25 * target_firms[name]["total_assets"] <= firm_total_assets <= 2 * target_firms[name]["total_assets"]:
				if name in result:
					if not firm in result[name]:
						result[name][firm] = 0
				else:
					result[name] = {}
					result[name][firm] = 0

	print "Target firms: {0}. Found matches for {1} target firms".format(len(target_firms),len(result))
	return result

def get_income_data():
	ws = wb2.active
	data = build_industry_groups()
	target_firms = build_target_firm_data()
	company_col = column_index_from_string('I') - 1
	date_col = column_index_from_string('B') - 1
	net_income_col = column_index_from_string('T') - 1
	for row in ws.iter_rows(row_offset=1):
		datecell = row[date_col].value
		date = int(str(datecell)[:4])
		firm = unicode(row[company_col].value).strip()

		net_income = row[net_income_col].value
		if not net_income:
			continue
		for target_firm in data:
			if firm in data[target_firm] and date == target_firms[target_firm]["pre_eventyear"]:
				data[target_firm][firm] = net_income
	return data



def get_match():
	ws = wb2.active
	data = get_income_data()
	target_firms = build_target_firm_data()

	# company_col = column_index_from_string('I') - 1
	# date_col = column_index_from_string('B') - 1
	# net_income_col = column_index_from_string('T') - 1
	result= {}
	for target_firm in data:
		income_diffs = {}
		for match in data[target_firm]:
			income_diffs[match] = (abs(target_firms[target_firm]["net_income"] - data[target_firm][match]))
		lowest_diff = min(income_diffs.values())
		for firm in income_diffs:
			if income_diffs[firm] == lowest_diff:
				result[target_firm] = firm
	print len(result)
	return result





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

if __name__ == "__main__":
	# print build_target_firm_data()
	# print build_industry_groups()
	# print change_entry(check_pre_event_year())
	print get_match()