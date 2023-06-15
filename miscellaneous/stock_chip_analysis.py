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
import copy
from datetime import datetime
# from pymongo import MongoClient
from collections import OrderedDict


class StockChipAnalysis(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\\Users\\%s\\Downloads" % os.getlogin()
	DEFAULT_SOURCE_FILENAME = "stock_chip_analysis"
	DEFAULT_SOURCE_FULL_FILENAME = "%s.xlsm" % DEFAULT_SOURCE_FILENAME
	DEFAULT_CONFIG_FOLDERPATH =  "C:\\Users\\%s" % os.getlogin()
	DEFAULT_DISPLAY_STOCK_LIST_FILENAME = "chip_analysis_stock_list.txt"
	# DEFAULT_REPORT_FILENAME = "chip_analysis_report.xlsx"
	DEFAULT_OUTPUT_RESULT_FILENAME = "output_result.txt"
	SHEET_METADATA_DICT = {
		# u"短線多空": {
		# 	"key_mode": 0, # 2504 國產
		# 	"data_start_column_index": 1,
		# },
		u"夏普值": {
			"key_mode": 0, # 2489 瑞軒
			"data_start_column_index": 1,
		},
		u"主法量率": {
			"key_mode": 0, # 2504 國產
			"data_start_column_index": 1,
		},
		u"六大買超": {
			"key_mode": 0, # 2504 國產
			"data_start_column_index": 1,
		},
		# u"大戶持股變化": {
		# 	"key_mode": 0, # 2504 國產
		# 	"data_start_column_index": 1,
		# },
		u"主力買超天數累計": {
			"key_mode": 0, # 2504 國產
			"data_start_column_index": 1,
		},
		# u"法人買超天數累計": {
		# 	"key_mode": 0, # 2504 國產
		# 	"data_start_column_index": 1,
		# },
		u"法人共同買超累計": {
			"key_mode": 1, # 1476
			"data_start_column_index": 2,
		},
		u"外資買超天數累計": {
			"key_mode": 0, # 2504 國產
			"data_start_column_index": 1,
		},
		u"投信買超天數累計": {
			"key_mode": 0, # 2504 國產
			"data_start_column_index": 1,
		},
		u"上市融資增加": {
			"key_mode": 2, # 4736  泰博
			"data_start_column_index": 1,
		},
		u"上櫃融資增加": {
			"key_mode": 2, # 4736  泰博
			"data_start_column_index": 1,
		},
	}
	ALL_SHEET_NAME_LIST = SHEET_METADATA_DICT.keys()
	DEFAULT_SHEET_NAME_LIST = [u"夏普值", u"主法量率", u"六大買超", u"主力買超天數累計", u"法人共同買超累計", u"外資買超天數累計", u"投信買超天數累計", u"上市融資增加", u"上櫃融資增加",]
	SHEET_SET_LIST = [
		[u"法人共同買超累計", u"主力買超天數累計", u"外資買超天數累計", u"投信買超天數累計",],
		[u"法人共同買超累計", u"外資買超天數累計", u"投信買超天數累計",],
		[u"外資買超天數累計", u"投信買超天數累計",],
	]
	DEFAULT_MIN_CONSECUTIVE_OVER_BUY_DAYS = 3
	DEFAULT_MAX_CONSECUTIVE_OVER_BUY_DAYS = 15
	CONSECUTIVE_OVER_BUY_DAYS_SHEETNAME_LIST = [u"主力買超天數累計", u"外資買超天數累計", u"投信買超天數累計",]
	CONSECUTIVE_OVER_BUY_DAYS_FIELDNAME_LIST = [u"主力買超累計天數", u"外資買超累計天數", u"投信買超累計天數",]
	DEFAULT_MINIMUM_VOLUME = 1000
	MINIMUM_VOLUME_SHEETNAME_LIST = [u"主力買超天數累計", u"外資買超天數累計", u"投信買超天數累計",]
	MINIMUM_VOLUME_FIELDNAME_LIST = [u"主力買超張數", u"外資累計買超張數", u"投信累計買超張數",]
	DEFAULT_MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_THRESHOLD = 10.0
	DEFAULT_MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_CONSECUTIVE_DAYS = 3
	MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_SHEETNAME = "主法量率"
	MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_FIELDNAME = "主法量率 D"

	SEARCH_RULE_DATASHEET_LIST = [
		["主法量率", "主力買超天數累計", "外資買超天數累計", "投信買超天數累計",],
		["主法量率", "主力買超天數累計", "外資買超天數累計",],
		["主法量率", "主力買超天數累計", "投信買超天數累計",],
		["主法量率", "主力買超天數累計",],
	]

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
				print ("%s: %s" % (errno.errorcode[exception.errno], os.strerror(exception.errno)))
				raise
			check_exist = False
		return check_exist


	@classmethod
	def __read_from_worksheet(cls, worksheet, sheet_metadata):
		# import pdb; pdb.set_trace()
		csv_data_dict = {}

		sheet_name = worksheet.name
		start_column_index = sheet_metadata["data_start_column_index"]
		title_list = ["商品",]
		for column_index in range(start_column_index, worksheet.ncols):
			title = worksheet.cell_value(0, column_index)
			title_list.append(title)

		for row_index in range(1, worksheet.nrows):
			data_list = []
			ignore_data = False
			stock_number = None
			product_name = None
			key_str = worksheet.cell_value(row_index, 0)
			# print "key_str: %s" % key_str
			if sheet_metadata["key_mode"] == 0:
				mobj = re.match("([\d]{4})\s(.+)", key_str)
				if mobj is None:
					raise ValueError("%s: Incorrect format0: %s" % (sheet_name, key_str))
				stock_number = mobj.group(1)
				product_name = mobj.group(2)
			elif sheet_metadata["key_mode"] == 1:
				# mobj = re.match("([\d]{4})\.TW", key_str)
				mobj = re.match("([\d]{4})", str(int(key_str)))
				if mobj is None:
					raise ValueError("%s: Incorrect format1: %s" % (sheet_name, key_str))
				stock_number = mobj.group(1)
				product_name = worksheet.cell_value(row_index, 1)
			elif sheet_metadata["key_mode"] == 2:
				mobj = re.match("([\d]{4})\s{2}(.+)", key_str)
				if mobj is None:
					ignore_data = True
				else:
					stock_number = mobj.group(1)
					product_name = mobj.group(2)
			else:
				raise ValueError("Unknown key mode: %d" % sheet_metadata["key_mode"])
			# if stock_number is None:
			#	raise RuntimeError("Fail to parse the stock number")
			if not ignore_data:
				data_list.append(product_name)
				for column_index in range(start_column_index, worksheet.ncols):
					data_list.append(worksheet.cell_value(row_index, column_index))
			# print "%d -- %s" % (row_index, stock_number)
			csv_data_dict[stock_number] = dict(zip(title_list, data_list))
		return csv_data_dict


	@classmethod
	def show_search_targets_list(cls):
		print("*****************************************")
		for index, search_rule_dataset in enumerate(cls.SEARCH_RULE_DATASHEET_LIST):
			print(" %d: %s" % (index, ",".join(search_rule_dataset)))
		print("*****************************************")


	def __init__(self, cfg):
		self.xcfg = {
			"source_folderpath": None,
			"source_filename": self.DEFAULT_SOURCE_FULL_FILENAME,
			"display_stock_list_filename": self.DEFAULT_DISPLAY_STOCK_LIST_FILENAME,
			"display_stock_list": None,
			"min_consecutive_over_buy_days": self.DEFAULT_MIN_CONSECUTIVE_OVER_BUY_DAYS,
			"max_consecutive_over_buy_days": self.DEFAULT_MAX_CONSECUTIVE_OVER_BUY_DAYS,
			"minimum_volume": self.DEFAULT_MINIMUM_VOLUME,
			"main_force_instuitional_investors_ratio_threshold": self.DEFAULT_MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_THRESHOLD,
			"main_force_instuitional_investors_ratio_consecutive_days": self.DEFAULT_MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_CONSECUTIVE_DAYS,
			"output_result_filename": self.DEFAULT_OUTPUT_RESULT_FILENAME,
			"output_result": False,
			"quiet": False,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["source_folderpath"] = self.DEFAULT_SOURCE_FOLDERPATH if self.xcfg["source_folderpath"] is None else self.xcfg["source_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FULL_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_filename"])
		# print ("__init__: %s" % self.xcfg["source_filepath"])
		self.xcfg["display_stock_list_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["display_stock_list_filename"])
		if self.xcfg["display_stock_list"] is not None:
			if type(self.xcfg["display_stock_list"]) is str:
				display_stock_list = []
				for display_stock in self.xcfg["display_stock_list"].split(","):
					display_stock_list.append(display_stock)
				self.xcfg["display_stock_list"] = display_stock_list

		# import pdb; pdb.set_trace()
		self.xcfg["output_result_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["output_result_filename"])

		self.filepath_dict = OrderedDict()
		self.filepath_dict["source"] = self.xcfg["source_filepath"]
		self.filepath_dict["display_stock_list"] = self.xcfg["display_stock_list_filepath"]
		self.filepath_dict["output_result"] = self.xcfg["output_result_filepath"]

		self.workbook = None
		self.output_result_file = None
		self.stdout_tmp = None


	def __enter__(self):
		# Open the workbook
		# self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		# if self.xcfg["output_result"]:
		# 	self.output_result_file = open(self.xcfg["output_result_filepath"], "w")
		return self


	def __exit__(self, type, msg, traceback):
		if self.output_result_file is not None:
			self.output_result_file.close()
			self.output_result_file = None
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


	def __get_workbook(self):
		if self.workbook is None:
			# import pdb; pdb.set_trace()
			self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
			# print ("__get_workbook: %s" % self.xcfg["source_filepath"])
		return self.workbook


	def __redirect_stdout2file(self):
# The output is now directed to the file
		if self.output_result_file is None:
			self.output_result_file = open(self.xcfg["output_result_filepath"], 'w')
# Store the current STDOUT object for later use
		self.stdout_tmp = sys.stdout
# Redirect STDOUT to the file
		sys.stdout = self.output_result_file


	def __redirect_file2stdout(self):
# Restore the original STDOUT
		sys.stdout = self.stdout_tmp


	def __read_sheet_data(self, sheet_name):
		# import pdb; pdb.set_trace()
		sheet_metadata = self.SHEET_METADATA_DICT[sheet_name]		
		# print (u"Read sheet: %s" % sheet_name)

		# assert self.workbook is not None, "self.workbook should NOT be None"
		worksheet = self.__get_workbook().sheet_by_name(sheet_name)
		# https://www.itread01.com/content/1549650266.html
		# print worksheet.name,worksheet.nrows,worksheet.ncols    #Sheet1 6 4
# The data
		csv_data_dict = self.__read_from_worksheet(worksheet, sheet_metadata)
# Filter the data if necessary
		if (self.xcfg["min_consecutive_over_buy_days"] is not None) or (self.xcfg["max_consecutive_over_buy_days"] is not None):
			try:
				sheet_index = self.CONSECUTIVE_OVER_BUY_DAYS_SHEETNAME_LIST.index(sheet_name)
				# import pdb; pdb.set_trace()
				field_name = self.CONSECUTIVE_OVER_BUY_DAYS_FIELDNAME_LIST[sheet_index]
				filter_func_ptr = None
				if (self.xcfg["min_consecutive_over_buy_days"] is not None) and (self.xcfg["max_consecutive_over_buy_days"] is not None):
					filter_func_ptr = lambda x: (self.xcfg["max_consecutive_over_buy_days"] >= int(x[1][field_name]) >= self.xcfg["min_consecutive_over_buy_days"])
				elif self.xcfg["min_consecutive_over_buy_days"] is not None:
					filter_func_ptr = lambda x: (int(x[1][field_name]) >= self.xcfg["min_consecutive_over_buy_days"])
				elif self.xcfg["max_consecutive_over_buy_days"] is not None:
					filter_func_ptr = lambda x: (self.xcfg["max_consecutive_over_buy_days"] >= int(x[1][field_name]))
				if filter_func_ptr is not None:
					csv_data_dict = dict(filter(filter_func_ptr, csv_data_dict.items()))
			except ValueError as e: 
				pass
		if self.xcfg["minimum_volume"] is not None:
			try:
				sheet_index = self.MINIMUM_VOLUME_SHEETNAME_LIST.index(sheet_name)
				# import pdb; pdb.set_trace()
				field_name = self.MINIMUM_VOLUME_FIELDNAME_LIST[sheet_index]
				csv_data_dict = dict(filter(lambda x: int(x[1][field_name]) >= self.xcfg["minimum_volume"], csv_data_dict.items()))
			except ValueError as e: 
				pass
		if self.xcfg["main_force_instuitional_investors_ratio_threshold"] is not None:
			if sheet_name == self.MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_SHEETNAME:
				csv_data_dict = dict(filter(lambda x: float(x[1][self.MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_FIELDNAME]) >= self.xcfg["main_force_instuitional_investors_ratio_threshold"], csv_data_dict.items()))
				if self.xcfg["main_force_instuitional_investors_ratio_consecutive_days"] is not None:
					def check_consecutive_days(x):
						# import pdb; pdb.set_trace()
						for index in range(1, self.xcfg["main_force_instuitional_investors_ratio_consecutive_days"]):
							field_name = "D-%d" % index
							if x[1][field_name] < self.xcfg["main_force_instuitional_investors_ratio_threshold"]:
								return False
						return True
					csv_data_dict = dict(filter(lambda x: check_consecutive_days(x), csv_data_dict.items()))
		return csv_data_dict


	def get_stock_chip_data(self, sheet_name_list=None):
		stock_chip_data_dict = {}
		if sheet_name_list is None:
			sheet_name_list = self.DEFAULT_SHEET_NAME_LIST
		for sheet_name in sheet_name_list:
			stock_chip_data_dict[sheet_name] = self.__read_sheet_data(sheet_name)
		# import pdb; pdb.set_trace()
		return stock_chip_data_dict


	def search_targets(self, stock_chip_data_dict=None, search_rule_index=0):
		if stock_chip_data_dict is None:
			stock_chip_data_dict = self.get_stock_chip_data()
		if search_rule_index < 0 or search_rule_index >= len(self.SEARCH_RULE_DATASHEET_LIST):
			raise ValueError("Unsupport search_rule_index: %d" % search_rule_index)
		search_rule_list = self.SEARCH_RULE_DATASHEET_LIST[search_rule_index]
		stock_set = set(stock_chip_data_dict[search_rule_list[0]].keys())
		for search_rule in search_rule_list[1:]:
			stock_set &= set(stock_chip_data_dict[search_rule].keys())
		stock_list = list(stock_set)

		if self.xcfg["output_result"]:
			self.__redirect_stdout2file()
		print("************** Search **************")
		search_rule_list_str = ", ".join(search_rule_list)
		print ("搜尋規則: " + search_rule_list_str )
		stock_name_list = [stock_chip_data_dict[u"主力買超天數累計"][stock]["商品"] for stock in stock_list]
		stock_list_str = ", ".join(map(lambda x: "%s[%s]" % (x[0], x[1]), zip(stock_list, stock_name_list)))
		print (stock_list_str + "\n")
		# import pdb; pdb.set_trace()
		sheet_name_list = ["夏普值", "主法量率", "六大買超",]
		for index, stock in enumerate(stock_list):
			search_rule_item_list = []
			for search_rule in search_rule_list[1:]:
				sheet_index = self.CONSECUTIVE_OVER_BUY_DAYS_SHEETNAME_LIST.index(search_rule)
				field_name = self.CONSECUTIVE_OVER_BUY_DAYS_FIELDNAME_LIST[sheet_index]
				sheet_data_dict = stock_chip_data_dict[search_rule]
				stock_sheet_data_dict = sheet_data_dict[stock]
				search_rule_item_list.append((field_name, str(int(stock_sheet_data_dict[field_name]))))
			print ("*** %s[%s] ***" % (stock, stock_name_list[index]))
			global_item_list = None
			for sheet_name in sheet_name_list:				
				sheet_data_dict = stock_chip_data_dict[sheet_name]
				if stock not in sheet_data_dict.keys():
					continue
				stock_sheet_data_dict = sheet_data_dict[stock]
				item_list = stock_sheet_data_dict.items()
				if global_item_list is None:
					global_item_list = []
					global_item_list.extend(filter(lambda x: x[0] in ["成交", "漲幅%", "漲跌幅",], item_list))
					global_item_list.extend(map(lambda x: (x[0], str(int(x[1]))), filter(lambda x: x[0] in ["成交量", "總量",], item_list)))
					global_item_list.extend(search_rule_item_list)
					print(" ==>" + " ".join(map(lambda x: "%s(%s)" % (x[0], x[1]), global_item_list)))
				item_list = filter(lambda x: x[0] not in ["商品", "成交", "漲幅%", "漲跌幅", "成交量", "總量",], item_list)
				if sheet_name in ["六大買超", "法人共同買超累計",]:
					print("  " + " ".join(map(lambda x: "%s(%d)" % (x[0], int(x[1])), item_list)))
				else:
					print("  " + " ".join(map(lambda x: "%s(%s)" % (x[0], x[1]), item_list)))
			print("\n")
		if self.xcfg["output_result"]:
			self.__redirect_file2stdout()


	def display_targets(self, stock_chip_data_dict=None):
		if self.xcfg["display_stock_list"] is None:
			self.__get_display_stock_list_from_file()
		if stock_chip_data_dict is None:
			stock_chip_data_dict = self.get_stock_chip_data()

		if self.xcfg["output_result"]:
			self.__redirect_stdout2file()
		print("************** Display **************")
		for display_stock in self.xcfg["display_stock_list"]:
			# print ("*** %s[%s] ***" % (display_stock, stock_name_list[index]))
			target_caption = None
			global_item_list = None
			need_new_line = False
			for sheet_name in self.DEFAULT_SHEET_NAME_LIST:
				sheet_data_dict = stock_chip_data_dict[sheet_name]
				if display_stock not in sheet_data_dict.keys():
					continue
				if not need_new_line:
					need_new_line = True
				stock_sheet_data_dict = sheet_data_dict[display_stock]
				if target_caption is None:
					target_caption = "*** %s[%s] ***" % (display_stock, stock_sheet_data_dict["商品"])
					print(target_caption)
				item_list = stock_sheet_data_dict.items()
				if global_item_list is None:
					global_item_list = []
					global_item_list.extend(filter(lambda x: x[0] in ["成交", "漲幅%", "漲跌幅",], item_list))
					global_item_list.extend(map(lambda x: (x[0], str(int(x[1]))), filter(lambda x: x[0] in ["成交量", "總量",], item_list)))
					print(" ==>" + " ".join(map(lambda x: "%s(%s)" % (x[0], x[1]), global_item_list)))
				item_list = filter(lambda x: x[0] not in ["商品", "成交", "漲幅%", "漲跌幅", "成交量", "總量",], item_list)
				if sheet_name in ["六大買超", "主力買超天數累計", "法人共同買超累計", "外資買超天數累計", "投信買超天數累計",]:
					print("  " + " ".join(map(lambda x: "%s(%d)" % (x[0], int(x[1])), item_list)))
				else:
					print("  " + " ".join(map(lambda x: "%s(%s)" % (x[0], x[1]), item_list)))
			if need_new_line: 
				print("\n")
		if self.xcfg["output_result"]:
			self.__redirect_file2stdout()


	def __get_display_stock_list_from_file(self):
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(self.xcfg['display_stock_list_filepath']):
			raise RuntimeError("The file[%s] does NOT exist" % self.xcfg['display_stock_list_filepath'])
		self.xcfg["display_stock_list"] = []
		with open(self.xcfg['display_stock_list_filepath'], 'r') as fp:
			for line in fp:
				self.xcfg["display_stock_list"].append(line.strip("\n"))


	def search_sheets(self, search_whole=False):
		if search_whole:
			if self.xcfg['stock_list'] is not None:
				raise RuntimeError("The stock list should be None")
		else:
			if self.xcfg['stock_list'] is None:
				raise RuntimeError("The stock list should NOT be None")
			self.xcfg['stock_list'] = self.xcfg['stock_list'].split(",")
		self.__search_stock_sheets()


	def print_filepath(self):
		print("************** File Path **************")
		for key, value in self.filepath_dict.items():
			print("%s: %s" % (key, value))


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
	'store_true' and 'store_false' - ?些是 'store_const' 分?用作存? True 和 False 值的特殊用例。
	另外，它?的默?值分?? False 和 True。例如:

	>>> parser = argparse.ArgumentParser()
	>>> parser.add_argument('--foo', action='store_true')
	>>> parser.add_argument('--bar', action='store_false')
	>>> parser.add_argument('--baz', action='store_false')
	'''
	parser.add_argument('-l', '--list_search_rule', required=False, action='store_true', help='List each search rule and exit.')
	parser.add_argument('-r', '--search_rule', required=False, help='The rule for selecting targets. Default: 0.')
	parser.add_argument('-s', '--search', required=False, action='store_true', help='Select targets based on the search rule.')
	parser.add_argument('-d', '--display', required=False, action='store_true', help='Display specific targets.')
	parser.add_argument('--display_stock_list', required=False, help='The list of specific targets to be displayed.')
	parser.add_argument('--print_filepath', required=False, action='store_true', help='Print the filepaths used in the process and exit.')
	parser.add_argument('-o', '--output_result', required=False, action='store_true', help='Output the result to the file instead of STDOUT.')
	parser.add_argument('--output_result_filename', required=False, action='store_true', help='The filename of outputing the result to the file instead of STDOUT.')
	args = parser.parse_args()

	if args.list_search_rule:
		StockChipAnalysis.show_search_targets_list()
		sys.exit(0)

	cfg = {}
	if args.display_stock_list:
		cfg['display_stock_list'] = args.display_stock_list
	if args.output_result:
		cfg['output_result'] = True
	if args.output_result_filename:
		cfg['output_result_filename'] = args.output_result_filename
	with StockChipAnalysis(cfg) as obj:
		if args.print_filepath:
			obj.print_filepath()
			sys.exit(0)
		if args.search:
			search_rule_index = int(args.search_rule) if args.search_rule else 0
			obj.search_targets(search_rule_index=search_rule_index)
		if args.display:
			obj.display_targets()
