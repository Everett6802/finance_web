#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import errno
'''
Question: How to Solve xlrd.biffh.XLRDError: Excel xlsx file; not supported ?
Answer : The latest version of xlrd(2.01) only supports .xls files. Installing the older version 1.2.0 to open .xlsx files.
'''
import xlrd
import xlsxwriter
import argparse
from collections import OrderedDict


class StockChipAnalysis(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\Users\Price\Downloads" # os.getcwd()
	DEFAULT_SOURCE_FILENAME = "stock_chip_analysis.xlsm"
	DEFAULT_CONFIG_FOLDERPATH =  "C:\Users\Price\source"
	DEFAULT_STOCK_LIST_FILENAME = "chip_analysis_stock_list.txt"
	DEFAULT_REPORT_FILENAME = "chip_analysis_report.xlsx"
	SHEET_METADATA_DICT = {
		u"即時指數": { # Dummy
			"is_dummy": True,
		},
		u"主要指數": { # Dummy
			"is_dummy": True,
		},
		u"外匯市場": { # Dummy
			"is_dummy": True,
		},
		u"商品市場": { # Dummy
			"is_dummy": True,
		},
		u"商品行情": { # Dummy
			"is_dummy": True,
		},
		u"資金流向": { # Dummy
			"is_dummy": True,
		},
		u"大盤籌碼多空勢力": { # Dummy
			"is_dummy": True,
		},
		u"焦點股": { 
			"key_mode": 0, # 1476.TW
		},
		u"法人共同買超累計": {
			"key_mode": 0, # 1476.TW
			"direction": "+",
		},
		u"主力買超天數累計": {
			"key_mode": 0, # 1476.TW
			"direction": "+",
		},
		u"法人買超天數累計": {
			"key_mode": 0, # 1476.TW
			"direction": "+",
		},
		u"外資買超天數累計": {
			"key_mode": 0, # 1476.TW
			"direction": "+",
		},
		u"投信買超天數累計": {
			"key_mode": 0, # 1476.TW
			"direction": "+",
		},
		u"外資買最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "+",
		},
		u"外資賣最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "-",
		},
		u"投信買最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "+",
		},
		u"投信賣最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "-",
		},
		u"主力買最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "+",
		},
		u"主力賣最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "-",
		},
		u"籌碼排行-買超金額": {
			"key_mode": 1, # 陽明(2609)
			"direction": "+",
		},
		u"籌碼排行-賣超金額": {
			"key_mode": 1, # 陽明(2609)
			"direction": "-",
		},
		u"買超異常": {
			"key_mode": 1, # 陽明(2609)
			"direction": "+",
		},
		u"賣超異常": {
			"key_mode": 1, # 陽明(2609)
			"direction": "-",
		},
	}
	DEFAULT_SHEET_NAME_LIST = [u"焦點股", u"法人共同買超累計", u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計", u"外資買最多股", u"外資賣最多股", u"投信買最多股", u"投信賣最多股", u"主力買最多股", u"主力賣最多股", u"籌碼排行-買超金額", u"籌碼排行-賣超金額", u"買超異常", u"賣超異常",]
	SHEET_SET_LIST = [
		[u"法人共同買超累計", u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計",],
		[u"法人共同買超累計", u"外資買超天數累計", u"投信買超天數累計",],
	]
	DEFAULT_CONSECUTIVE_OVER_BUY_DAYS = 3
	CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET = [u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計",]
	CHECK_CONSECUTIVE_OVER_BUY_DAYS_FIELD_NAME_KEY = u"買超累計天數"
	@classmethod
	def __is_string(cls, value):
		is_string = False
		try:
			int(value)
		except ValueError:
			is_string = True
		return is_string


	@classmethod
	def __check_file_exist(cls, filepath):
		check_exist = True
		try:
			os.stat(filepath)
		except OSError as exception:
			if exception.errno != errno.ENOENT:
				print "%s: %s" % (errno.errorcode[exception.errno], os.strerror(exception.errno))
				raise
			check_exist = False
		return check_exist


	@classmethod
	def read_stock_list_from_file(cls, stock_list_filepath):
		# import pdb; pdb.set_trace()
		if not cls.__check_file_exist(stock_list_filepath):
			raise RuntimeError("The file[%s] does NOT exist" % stock_list_filepath)
		stock_list = []
		with open(stock_list_filepath, 'r') as fp:
			for line in fp:
				stock_list.append(line.strip("\n"))
		return stock_list


	@classmethod
	def list_sheet_set(cls):
		for index, sheet_set in enumerate(cls.SHEET_SET_LIST):
			print "%d: %s" % (index, ",".join(sheet_set))


	def __init__(self, cfg):
		self.xcfg = {
			"show_detail": False,
			"generate_report": False,
			"source_filename": self.DEFAULT_SOURCE_FILENAME,
			"stock_list_filename": self.DEFAULT_STOCK_LIST_FILENAME,
			"report_filename": self.DEFAULT_REPORT_FILENAME,
			"stock_list": None,
			"sheet_name_list": None,
			"stock_set_category": -1,
			"consecutive_over_buy_days": self.DEFAULT_CONSECUTIVE_OVER_BUY_DAYS,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["source_filepath"] = os.path.join(self.DEFAULT_SOURCE_FOLDERPATH, self.xcfg["source_filename"])
		self.xcfg["stock_list_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["stock_list_filename"])
		self.xcfg["report_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["report_filename"])
		if self.xcfg["generate_report"]:
			if not self.xcfg["show_detail"]:
				print "WARNING: The 'show_detail' parameter is enabled while the 'generate_report' one is true"
				self.xcfg["show_detail"] = True

		if self.xcfg["stock_set_category"] != -1:
			if self.xcfg["sheet_name_list"] is not None:
				print "WARNING: The 'stock_set_category' setting overwrite the 'sheet_name_list' one"
			self.xcfg["sheet_name_list"] = self.SHEET_SET_LIST[self.xcfg["stock_set_category"]]

		self.workbook = None
		self.output_workbook = None
		self.sheet_title_bar_dict = {}


	def __enter__(self):
		# Open the workbook
		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		if self.xcfg["generate_report"]:
			self.output_workbook = xlsxwriter.Workbook(self.xcfg["report_filepath"])
		return self


	def __exit__(self, type, msg, traceback):
		if self.output_workbook is not None:
			self.output_workbook.close()
			del self.output_workbook
			self.output_workbook = None
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


	def __read_sheet_title_bar(self, sheet_name):
		# import pdb; pdb.set_trace()
		if not self.sheet_title_bar_dict.has_key(sheet_name):
			sheet_metadata = self.SHEET_METADATA_DICT[sheet_name]
			worksheet = self.workbook.sheet_by_name(sheet_name)
			title_bar_list = [u"商品",]
			column_start_index = None
			if sheet_metadata["key_mode"] == 0:
				column_start_index = 2
			elif sheet_metadata["key_mode"] == 1:
				column_start_index = 1
			else:
				raise ValueError("Unknown key mode: %d" % sheet_metadata["key_mode"]) 
			for column_index in range(column_start_index, worksheet.ncols):
				title_bar_list.append(worksheet.cell_value(0, column_index))
			self.sheet_title_bar_dict[sheet_name] = title_bar_list
		return self.sheet_title_bar_dict[sheet_name]


	def __read_sheet_data(self, sheet_name):
		sheet_metadata = self.SHEET_METADATA_DICT[sheet_name]
		# print u"Read sheet: %s" % sheet_metadata["description"].decode("utf8")
		assert self.workbook is not None, "self.workbook should NOT be None"
		worksheet = self.workbook.sheet_by_name(sheet_name)
		# https://www.itread01.com/content/1549650266.html
		# print worksheet.name,worksheet.nrows,worksheet.ncols    #Sheet1 6 4
		data_dict = {}
		row_index = 1
		while True:
			try:
				key_str = worksheet.cell_value(row_index, 0)
			except IndexError:
				# print "Total rows: %d" % row_index
				break
			stock_number = None
			if sheet_metadata["key_mode"] == 0:
				mobj = re.match("([\d]{4})\.TW", key_str)
				stock_number = mobj.group(1)
				data_dict[stock_number] = []
			elif sheet_metadata["key_mode"] == 1:
				mobj = re.match("(.+)\(([\d]{4}[\d]?[\w]?)\)", key_str)
				stock_number = mobj.group(2)
				data_dict[stock_number] = [mobj.group(1),]
			else:
				raise ValueError("Unknown key mode: %d" % sheet_metadata["key_mode"])
			if stock_number is None:
				raise RuntimeError("Fail to parse the stock number")
			for column_index in range(1, worksheet.ncols):
				data_dict[stock_number].append(worksheet.cell_value(row_index, column_index))
			row_index += 1
			# print "%d -- %s" % (row_index, stock_number)
		if self.xcfg["consecutive_over_buy_days"] > 0:
			if sheet_name in self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET:
				data_dict = self.__filter_consecutive_over_buy_days(sheet_name, data_dict)
		return data_dict


	def __filter_consecutive_over_buy_days(self, sheet_name, data_dict):
		title_bar_list = self.__read_sheet_title_bar(sheet_name)
		found = False
		found_index = -1
		for index, title_bar in enumerate(title_bar_list):
			if re.search(self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_FIELD_NAME_KEY, title_bar):
				found = True
				found_index = index
				break
		assert found, "The %s is NOT found" % self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_FIELD_NAME_KEY
		# import pdb; pdb.set_trace()
		return dict(filter(lambda x: int(x[1][found_index]) >= self.xcfg["consecutive_over_buy_days"], data_dict.items()))


	def __collect_sheet_all_data(self, sheet_data_func_ptr=None):
		sheet_data_collection_dict = {}
		if self.xcfg["sheet_name_list"] is None:
			self.xcfg["sheet_name_list"] = self.DEFAULT_SHEET_NAME_LIST
		for sheet_name in self.xcfg["sheet_name_list"]:
			data_dict = self.__read_sheet_data(sheet_name)
			for data_key, data_value in data_dict.items():
				if not sheet_data_collection_dict.has_key(data_key):
					sheet_data_collection_dict[data_key] = {}
				if sheet_data_func_ptr is not None:
					 data_value = sheet_data_func_ptr(data_value)
				sheet_data_collection_dict[data_key][sheet_name] = data_value
		return sheet_data_collection_dict


	def __collect_sheet_data(self, sheet_data_func_ptr=None):
		if self.xcfg["stock_list"] is None:
			return self.__collect_sheet_all_data(sheet_data_func_ptr)
		sheet_data_collection_dict = {}
		if self.xcfg["sheet_name_list"] is None:
			self.xcfg["sheet_name_list"] = self.DEFAULT_SHEET_NAME_LIST
		for sheet_name in self.xcfg["sheet_name_list"]:
			data_dict = self.__read_sheet_data(sheet_name)
			for stock in self.xcfg["stock_list"]:
				if not data_dict.has_key(stock):
					continue
				stock_data = data_dict[stock]
				if not sheet_data_collection_dict.has_key(stock):
					sheet_data_collection_dict[stock] = {}
				if sheet_data_func_ptr is not None:
					 stock_data = sheet_data_func_ptr(stock_data)
				sheet_data_collection_dict[stock][sheet_name] = stock_data
		return sheet_data_collection_dict


	# def __find_sheet_occurrence(self, ignore_sheet_func_ptr=None, sheet_data_func_ptr=None):
	# 	stock_number_sheet_dict = {}
	# 	stock_number_extra_dict = {}
	# 	# import pdb; pdb.set_trace()
	# 	if self.xcfg["sheet_name_list"] is None:
	# 		self.xcfg["sheet_name_list"] = self.DEFAULT_SHEET_NAME_LIST
	# 	for sheet_index in self.xcfg["sheet_name_list"]:
	# 		if ignore_sheet_func_ptr is not None and ignore_sheet_func_ptr(sheet_index):
	# 			continue
	# 		data_dict = self.__read_sheet_data(sheet_index)
	# 		for stock_number, stock_data in data_dict.items():
	# 			if stock_number_sheet_dict.has_key(stock_number):
	# 				# stock_number_sheet_dict[stock_number] = stock_number_sheet_dict[stock_number] + 1
	# 				stock_number_sheet_dict[stock_number].append(sheet_index)					
	# 			else:
	# 				# stock_number_sheet_dict[stock_number] = 1
	# 				stock_number_sheet_dict[stock_number] = [sheet_index,]
	# 				if sheet_data_func_ptr is not None:
	# 					stock_number_extra_dict[stock_number] = sheet_data_func_ptr(stock_data)
	# 	return stock_number_sheet_dict, stock_number_extra_dict


	def __search_stock_sheets(self):
		# import pdb; pdb.set_trace()
		sheet_data_func_ptr = (lambda x: x) if self.xcfg["show_detail"] else (lambda x: x[0])
		sheet_data_collection_dict = self.__collect_sheet_data(sheet_data_func_ptr)
		if self.xcfg["stock_list"] is None:
			self.xcfg["stock_list"] = sheet_data_collection_dict.keys()
		no_data = True

		output_overview_worksheet = None
		output_overview_row = 0
		if self.xcfg["generate_report"]:
			output_overview_worksheet = self.output_workbook.add_worksheet("Overview")
					
		for stock_number in self.xcfg["stock_list"]:
			if not sheet_data_collection_dict.has_key(stock_number):
				continue
			# if re.search("6741", stock_number):
			# 	import pdb; pdb.set_trace()
			no_data = False
			stock_sheet_data_collection_dict = sheet_data_collection_dict[stock_number]
			if self.xcfg["show_detail"]:
				stock_name = stock_sheet_data_collection_dict.values()[0][0]
				print "=== %s(%s) ===" % (stock_number, stock_name)
				if self.xcfg["generate_report"]:
# For overview sheet
					output_overview_worksheet.write(output_overview_row, 0,  "%s(%s)" % (stock_number, stock_name))
					for output_overview_col, sheet_name in enumerate(stock_sheet_data_collection_dict.keys()):
						output_overview_worksheet.write(output_overview_row + 1, output_overview_col,  sheet_name)
					output_overview_row += 3
# For detailed sheet
					try:
						worksheet = self.output_workbook.add_worksheet("%s(%s)" % (stock_number, stock_name))
					except xlsxwriter.exceptions.InvalidWorksheetName:
						import pdb; pdb.set_trace()
						if re.match("6741", stock_number):
							worksheet = self.output_workbook.add_worksheet("%s(%s)" % (stock_number, stock_name.replace("*","")))
					output_row = 0
				for sheet_name, sheet_data_list in stock_sheet_data_collection_dict.items():
					sheet_title_bar_list = self.__read_sheet_title_bar(sheet_name)
					sheet_data_list_len = len(sheet_data_list)
					sheet_title_bar_list_len = len(sheet_title_bar_list)
					assert sheet_data_list_len == sheet_title_bar_list_len, "The list lengths are NOT identical, sheet_data_list_len: %d, sheet_title_bar_list_len: %d" % (sheet_data_list_len, sheet_title_bar_list_len)
					print "* %s" % sheet_name
					print "%s" % ",".join(["%s[%s]" % elem for elem in zip(sheet_title_bar_list[1:], sheet_data_list[1:])])
					if self.xcfg["generate_report"]:
# For detailed sheet
						worksheet.write(output_row, 0,  sheet_name)
						for output_col, output_data in enumerate(zip(sheet_title_bar_list[1:], sheet_data_list[1:])):
							sheet_title_bar, sheet_data = output_data
							worksheet.write(output_row + 1, output_col,  sheet_title_bar)
							worksheet.write(output_row + 2, output_col,  sheet_data)
						output_row += 4
			else:
				stock_name = stock_sheet_data_collection_dict.values()[0]
				print "=== %s(%s) ===" % (stock_number, stock_name)
				print "%s" % (u",".join([stock_sheet_data_key for stock_sheet_data_key in stock_sheet_data_collection_dict.keys()]))
		if no_data: print "*** No Data ***"	
		if self.xcfg["generate_report"]:
			if no_data:
				worksheet = workbook.add_worksheet("NoData")


	def search_sheets_from_file(self):
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(self.xcfg['stock_list_filepath']):
			raise RuntimeError("The file[%s] does NOT exist" % self.xcfg['stock_list_filepath'])
		self.xcfg["stock_list"] = []
		with open(self.xcfg['stock_list_filepath'], 'r') as fp:
			for line in fp:
				self.xcfg["stock_list"].append(line.strip("\n"))
		self.__search_stock_sheets()


	def search_sheets(self, search_whole=False):
		if search_whole:
			if self.xcfg['stock_list'] is not None:
				raise RuntimeError("The stock list should be None")
		else:
			if self.xcfg['stock_list'] is None:
				raise RuntimeError("The stock list should NOT be None")
			self.xcfg['stock_list'] = self.xcfg['stock_list'].split(",")
		self.__search_stock_sheets()


	# def search_buy(self):
	# 	# import pdb; pdb.set_trace()
	# 	sheet_occurrence_dict, sheet_occurrence_extra_dict = self.__find_sheet_occurrence(lambda x: self.SHEET_METADATA_LIST[x]["direction"] == '-', lambda x: x[0])
	# 	filtered_sheet_occurrence_dict = dict(filter(lambda x: len(x[1]) >= self.xcfg["buy_sheet_threshold"], sheet_occurrence_dict.items()))
	# 	filtered_sheet_occurrence_ordereddict = OrderedDict(sorted(filtered_sheet_occurrence_dict.items(), key=lambda x: x[1]))
	# 	for stock_number, sheet_name_list in filtered_sheet_occurrence_ordereddict.items():
	# 		print "=== %s(%s) ===" % (stock_number, sheet_occurrence_extra_dict[stock_number])
	# 		print "%s" % (u",".join([self.SHEET_METADATA_LIST[index]["description"] for index in sheet_name_list]))


	@property
	def StockList(self):
		return self.xcfg["stock_list"]


	@StockList.setter
	def StockList(self, stock_list):
		self.xcfg["stock_list"] = stock_list


if __name__ == "__main__":
	
	parser = argparse.ArgumentParser(description='Print help')
	'''
	參數基本上分兩種，一種是位置參數 (positional argument)，另一種就是選擇性參數 (optional argument)
	* example2.py
	parser.add_argument("pos1", help="positional argument 1")
	parser.add_argument("-o", "--optional-arg", help="optional argument", dest="opt", default="default")

	# python example2.py hello -o world 
	positional arg: hello
	optional arg: world
	'''
# How to add option without any argument? use action='store_true'
	'''
	'store_true' and 'store_false' - 这些是 'store_const' 分别用作存储 True 和 False 值的特殊用例。
	另外，它们的默认值分别为 False 和 True。例如:

	>>> parser = argparse.ArgumentParser()
	>>> parser.add_argument('--foo', action='store_true')
	>>> parser.add_argument('--bar', action='store_false')
	>>> parser.add_argument('--baz', action='store_false')
	'''
	parser.add_argument('-e', '--list_analysis_method', required=False, action='store_true', help='List each analysis method and exit')
	parser.add_argument('-i', '--list_stock_set_category', required=False, action='store_true', help='List each stock set and exit')
	parser.add_argument('-m', '--analysis_method', required=False, help='The method for chip analysis. Default: 0')	
	parser.add_argument('-d', '--show_detail', required=False, action='store_true', help='Show detailed data for each stock')
	parser.add_argument('-g', '--generate_report', required=False, action='store_true', help='Generate the report of the detailed data for each stock to the XLS file.')
	parser.add_argument('-r', '--report_filename', required=False, help='The filename of chip analysis report')
	parser.add_argument('-t', '--stock_list_filename', required=False, help='The filename of stock list for chip analysis')
	parser.add_argument('-l', '--stock_list', required=False, help='The list string of stock list for chip analysis. Ex: 2330,2317,2454,2308')
	parser.add_argument('-s', '--source_filename', required=False, help='The filename of chip analysis data source')
	parser.add_argument('-c', '--stock_set_category', required=False, help='The category for stock set. Default: 0')	
	args = parser.parse_args()

	if args.list_analysis_method:
		help_str_list = [
			"Search sheet for the specific stocks from the file",
			"Search sheet for the specific stocks",
			"Search sheet for the whole stocks",
		]
		help_str_list_len = len(help_str_list)
		print "************ Analysis Method ************"
		for index, help_str in enumerate(help_str_list):
			print "%d  %s" % (index, help_str)
		print "*****************************************"
		sys.exit(0)
	if args.list_stock_set_category:
		StockChipAnalysis.list_sheet_set()
		sys.exit(0)

	# import pdb; pdb.set_trace()
	cfg = {}
	cfg['analysis_method'] = int(args.analysis_method) if args.analysis_method is not None else 0
	if args.show_detail: cfg['show_detail'] = True
	if args.generate_report: cfg['generate_report'] = True
	if args.report_filename is not None: cfg['report_filename'] = args.report_filename
	if args.stock_list_filename is not None: cfg['stock_list_filename'] = args.stock_list_filename
	if args.stock_list is not None: cfg['stock_list'] = args.stock_list
	if args.source_filename is not None: cfg['source_filename'] = args.source_filename
	cfg['stock_set_category'] = int(args.stock_set_category) if args.stock_set_category is not None else -1
		
	# import pdb; pdb.set_trace()
	with StockChipAnalysis(cfg) as obj:
		if cfg['analysis_method'] == 0:
			obj.search_sheets_from_file()
		elif cfg['analysis_method'] == 1:
			obj.search_sheets()
		elif cfg['analysis_method'] == 2:
			obj.search_sheets(True)
		else:
			raise ValueError("Incorrect Analysis Method Index")
