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
	DEFAULT_STOCK_LIST_FILENAME = "chip_analysis_stock_list.txt"
	DEFAULT_REPORT_FILENAME = "chip_analysis_report.xlsx"
	DEFAULT_SEARCH_RESULT_FILENAME = "search_result_stock_list.txt"
	SHEET_METADATA_DICT = {
		# u"短線多空": {
		# 	"key_mode": 0, # 2504 國產
		# 	"data_start_column_index": 1,
		# },
		u"夏普值": {
			"key_mode": 1, # 1476
			"data_start_column_index": 2,
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

	DEFAULT_DB_NAME = "StockChipAnalysis"
	DEFAULT_DB_USERNAME = "root"
	DEFAULT_DB_PASSWORD = "lab4man1"
	DEFAULT_DB_DATE_STRING_FORMAT = "%Y-%m-%d"

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


	def __init__(self, cfg):
		self.xcfg = {
			"show_detail": False,
			"generate_report": False,
			"source_folderpath": None,
			"source_filename": self.DEFAULT_SOURCE_FULL_FILENAME,
			"stock_list_filename": self.DEFAULT_STOCK_LIST_FILENAME,
			"report_filename": self.DEFAULT_REPORT_FILENAME,
			"stock_list": None,
			"sheet_name_list": None,
			"sheet_set_category": -1,
			"min_consecutive_over_buy_days": self.DEFAULT_MIN_CONSECUTIVE_OVER_BUY_DAYS,
			"max_consecutive_over_buy_days": self.DEFAULT_MAX_CONSECUTIVE_OVER_BUY_DAYS,
			"minimum_volume": self.DEFAULT_MINIMUM_VOLUME,
			"main_force_instuitional_investors_ratio_threshold": self.DEFAULT_MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_THRESHOLD,
			"main_force_instuitional_investors_ratio_consecutive_days": self.DEFAULT_MAIN_FORCE_INSTUITIONAL_INVESTORS_RATIO_CONSECUTIVE_DAYS,
			"need_all_sheet": False,
			"search_history": False,
			"search_result_filename": self.DEFAULT_SEARCH_RESULT_FILENAME,
			"output_search_result": False,
			"quiet": False,
			"sort": False,
			"sort_limit": None,
			"db_enable": False,
			"db_host": "localhost",
			"db_name": self.DEFAULT_DB_NAME,
			"db_username": self.DEFAULT_DB_USERNAME,
			"db_password": self.DEFAULT_DB_PASSWORD,
			"database_date": None,
			"database_date_range": None,
			"database_date_range_start": None,
			"database_date_range_end": None,
			"database_all_date_range": False,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["source_folderpath"] = self.DEFAULT_SOURCE_FOLDERPATH if self.xcfg["source_folderpath"] is None else self.xcfg["source_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FULL_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_filename"])
		# print ("__init__: %s" % self.xcfg["source_filepath"])
		self.xcfg["stock_list_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["stock_list_filename"])
		self.xcfg["report_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["report_filename"])
		self.xcfg["search_result_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["search_result_filename"])
		if self.xcfg["generate_report"]:
			if not self.xcfg["show_detail"]:
				print ("WARNING: The 'show_detail' parameter is enabled while the 'generate_report' one is true")
				self.xcfg["show_detail"] = True

		if self.xcfg["sheet_set_category"] != -1:
			if self.xcfg["sheet_name_list"] is not None:
				print ("WARNING: The 'sheet_set_category' setting overwrite the 'sheet_name_list' one")
			self.xcfg["sheet_name_list"] = self.SHEET_SET_LIST[self.xcfg["sheet_set_category"]]
		if self.xcfg["database_date"] is not None:
			if re.match("20[\d]{2}-[\d]{2}-[\d]{2}", self.xcfg["database_date"]) is None:
				raise ValueError("Incorrect date format: %s" % self.xcfg["database_date"])
		check_date_input = (self.xcfg["database_date"] is not None) and (self.xcfg["database_date_range"] is not None)
		assert not check_date_input, "database_date/database_date_range can NOT be set simultaneously"
		if self.xcfg["database_date_range"] is not None:
			elem_list = self.xcfg["database_date_range"].split(",")
			if len(elem_list) != 2:
				raise ValueError("Incorrect date range format: %s" % self.xcfg["database_date_range"])
			if len(elem_list[0]) != 0:
				if re.match("20[\d]{2}-[\d]{2}-[\d]{2}", elem_list[0]) is None:
					raise ValueError("Incorrect start date format: %s" % elem_list[0])
				self.xcfg["database_date_range_start"] = elem_list[0]
			if len(elem_list[1]) != 0:
				if re.match("20[\d]{2}-[\d]{2}-[\d]{2}", elem_list[1]) is None:
					raise ValueError("Incorrect end date format: %s" % elem_list[1])
				self.xcfg["database_date_range_end"] = elem_list[1]

		self.workbook = None
		self.report_workbook = None
		self.search_result_txtfile = None
		self.sheet_title_bar_dict = {}
		self.db_client = None
		self.db_handle = None


	def __enter__(self):
		# Open the workbook
		# self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		if self.xcfg["output_search_result"]:
			self.search_result_txtfile = open(self.xcfg["search_result_filepath"], "w")

# mongodb://root:lab4man1@localhost:27017/StockChipAnalysis
		# db_url = 'mongodb://%s:%s@%s:27017' % (self.xcfg["db_username"], self.xcfg["db_password"], self.xcfg["db_host"])
		if self.xcfg["db_enable"]:
			from pymongo import MongoClient
			db_url = 'mongodb://%s:%s@%s:27017/%s' % (self.xcfg["db_username"], self.xcfg["db_password"], self.xcfg["db_host"], self.xcfg["db_name"])
			# print ("DB URL: %s" % db_url)
			self.db_client = MongoClient(db_url)
# Database (Database -> Collection -> Document)
			self.db_handle = self.db_client[self.xcfg["db_name"]]
		return self


	def __exit__(self, type, msg, traceback):
		if self.db_client is not None:
			self.db_handle = None
			self.db_client.close()
			self.db_client = None
		if self.search_result_txtfile is not None:
			self.search_result_txtfile.close()
			self.search_result_txtfile = None
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


	def __print_string(self, outpug_str):
		if self.xcfg["quiet"]: return
		print (outpug_str)


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


	def search_targets(self, stock_chip_data_dict, search_rule_index=1):
		stock_set1 = set(stock_chip_data_dict[u"主法量率"].keys())
		stock_set2 = set(stock_chip_data_dict[u"主力買超天數累計"].keys())
		stock_set3 = set(stock_chip_data_dict[u"外資買超天數累計"].keys())
		stock_set4 = set(stock_chip_data_dict[u"投信買超天數累計"].keys())
# https://blog.csdn.net/qq_37195276/article/details/79467917
# & != and ; | != or
# python中&、|代表的是位運算符，and、or代表的是邏輯運算符
		search_rule_list = None
		stock_list = None
		if search_rule_index == 0:
			search_rule_list = ["主法量率", "主力買超天數累計", "外資買超天數累計", "投信買超天數累計",]
			stock_list = list(stock_set1 & stock_set2 & stock_set3 & stock_set4)
		elif search_rule_index == 1:
			search_rule_list = ["主法量率", "主力買超天數累計", "外資買超天數累計",]
			stock_list = list(stock_set1 & stock_set2 & stock_set3)
		elif search_rule_index == 2:
			search_rule_list = ["主法量率", "主力買超天數累計", "投信買超天數累計",]
			stock_list = list(stock_set1 & stock_set2 & stock_set4)
		elif search_rule_index == 3:
			search_rule_list = ["主法量率", "主力買超天數累計",]
			stock_list = list(stock_set1 & stock_set2)
		else:
			raise ValueError("Unsupport search_rule_index: %d" % search_rule_index)

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
	# parser.add_argument('--list_analysis_method', required=False, action='store_true', help='List each analysis method and exit')
	# parser.add_argument('--list_sheet_set_category', required=False, action='store_true', help='List each stock set and exit')
	# parser.add_argument('--update_database', required=False, action='store_true', help='Update database and exit')
	# parser.add_argument('--update_database_multiple', required=False, action='store_true', help='Update database from multiple XLS files and exit. Caution: The format of the XLS filename: {0}@20YY-mm-DD. Ex: {0}_2022-07-29'.format(StockChipAnalysis.DEFAULT_SOURCE_FILENAME))
	# parser.add_argument('--find_database', required=False, action='store_true', help='Find database and exit')
	# parser.add_argument('--delete_database', required=False, action='store_true', help='Delete database and exit')
	# parser.add_argument('--list_database_date', required=False, action='store_true', help='List database date and exit')
	# parser.add_argument('--database_date', required=False, help='The date of the data in the database. Ex: 2022-05-18. Caution: Update/Find/Delete Database')
	# parser.add_argument('--database_date_range', required=False, help='The date range of the data in the database. Format: start_date,end_date. Ex: (1) 2022-05-18,2022-05-30 ; (2) 2022-05-18, ; (3) ,2022-05-30. Caution: Find/Delete Database')
	# parser.add_argument('--database_all_date_range', required=False, action='store_true', help='The all date range of the data in the database')
	# parser.add_argument('--create_report_by_sheet_set_category', required=False, help='Create a report by certain a sheet set category and exit')
	# parser.add_argument('--check_data_source_difference', required=False, help='Check data source difference on a specific day and exit')
	# parser.add_argument('-m', '--analysis_method', required=False, help='The method for chip analysis. Default: 0')	
	# parser.add_argument('-d', '--show_detail', required=False, action='store_true', help='Show detailed data for each stock')
	# parser.add_argument('-g', '--generate_report', required=False, action='store_true', help='Generate the report of the detailed data for each stock to the XLS file.')
	# parser.add_argument('-r', '--report_filename', required=False, help='The filename of chip analysis report')
	# parser.add_argument('-t', '--stock_list_filename', required=False, help='The filename of stock list for chip analysis')
	# parser.add_argument('-l', '--stock_list', required=False, help='The list string of stock list for chip analysis. Ex: 2330,2317,2454,2308')
	# parser.add_argument('--source_folderpath', required=False, help='Update database from the XLS files in the designated folder path. Ex: %s' % StockChipAnalysis.DEFAULT_SOURCE_FOLDERPATH)
	# parser.add_argument('--source_filename', required=False, help='The filename of chip analysis data source')
	# parser.add_argument('-c', '--sheet_set_category', required=False, help='The category for sheet set. Default: 0')	
	# parser.add_argument('-n', '--need_all_sheet', required=False, action='store_true', help='The stock should be found in all sheets in the sheet name list')
	# parser.add_argument('--search_history', required=False, action='store_true', help='The data source is from the database, otherwise from the excel file')
	# parser.add_argument('-a', '--search_result_filename', required=False, help='The filename of stock list for search result')
	# parser.add_argument('-o', '--output_search_result', required=False, action='store_true', help='Ouput the search result')
	# parser.add_argument('-q', '--quiet', required=False, action='store_true', help="Don't print string on the screen")
	# parser.add_argument('-s', '--sort', required=False, action='store_true', help="Show the data in order")
	# parser.add_argument('-f', '--sort_limit', required=False, help="Limit the sorted data")
	# # parser.add_argument('-f', '--update_datebase', required=False, help="Limit the sorted data")
	args = parser.parse_args()

	# if args.list_analysis_method:
	# 	help_str_list = [
	# 		"Search data for the specific stocks from the file",
	# 		"Search data for the specific stocks",
	# 		"Search data for the whole stocks",
	# 		"Evaluate the TAIEX bull or bear",
	# 	]
	# 	help_str_list_len = len(help_str_list)
	# 	print ("************ Analysis Method ************")
	# 	for index, help_str in enumerate(help_str_list):
	# 		print ("%d  %s" % (index, help_str))
	# 	print ("*****************************************")
	# 	sys.exit(0)
	# if args.list_sheet_set_category:
	# 	StockChipAnalysis.list_sheet_set()
	# 	sys.exit(0)
	# if args.check_data_source_difference:
	# 	StockChipAnalysis.check_data_source_difference(args.check_data_source_difference)
	# 	sys.exit(0)
	# if args.create_report_by_sheet_set_category:
	# 	search_result_filename = "tmp1.txt"
	# 	cfg_step1 = {
	# 		"sheet_set_category": int(args.create_report_by_sheet_set_category),
	# 		"need_all_sheet": True,
	# 		"search_result_filename": search_result_filename,
	# 		"output_search_result": True,
	# 		"quiet": True,
	# 	}
	# 	with StockChipAnalysis(cfg_step1) as obj_step1: 
	# 		obj_step1.search_sheets(True)
	# 	cfg_step2 = {
	# 		"generate_report": True,
	# 		"stock_list_filename": search_result_filename,
	# 		"quiet": True,
	# 	}
	# 	with StockChipAnalysis(cfg_step2) as obj_step2: 
	# 		obj_step2.search_sheets_from_file()
	# 		os.remove(obj_step2.StockListFilepath)
	# 	sys.exit(0)

	# # import pdb; pdb.set_trace()
	# cfg = {}
	# cfg['analysis_method'] = int(args.analysis_method) if args.analysis_method is not None else 0
	# if args.database_date is not None: cfg['database_date'] = args.database_date
	# if args.database_date_range is not None: cfg['database_date_range'] = args.database_date_range
	# if args.database_all_date_range: cfg['database_all_date_range'] = True
	# if args.show_detail: cfg['show_detail'] = True
	# if args.generate_report: cfg['generate_report'] = True
	# if args.report_filename is not None: cfg['report_filename'] = args.report_filename
	# if args.stock_list_filename is not None: cfg['stock_list_filename'] = args.stock_list_filename
	# if args.stock_list is not None: cfg['stock_list'] = args.stock_list
	# if args.source_folderpath is not None: cfg['source_folderpath'] = args.source_folderpath
	# if args.source_filename is not None: cfg['source_filename'] = args.source_filename
	# cfg['sheet_set_category'] = int(args.sheet_set_category) if args.sheet_set_category is not None else -1
	# if args.need_all_sheet: cfg['need_all_sheet'] = True
	# if args.report_filename is not None: cfg['report_filename'] = args.report_filename
	# if args.search_history: cfg['search_history'] = True
	# if args.search_result_filename is not None: cfg['search_result_filename'] = args.search_result_filename
	# if args.output_search_result: cfg['output_search_result'] = True
	# if args.quiet: cfg['quiet'] = True
	# if args.sort: cfg['sort'] = True
	# if args.sort_limit is not None: cfg['sort_limit'] = int(args.sort_limit)

	# # import pdb; pdb.set_trace()
	# with StockChipAnalysis(cfg) as obj:
	# 	if args.update_database:
	# 		 obj.update_database()
	# 	elif args.update_database_multiple:
	# 		fliename_dict = {}
	# 		pattern = "%s.xlsm|%s@(20[\d]{2}-[\d]{2}-[\d]{2}).xlsm" % (StockChipAnalysis.DEFAULT_SOURCE_FILENAME, StockChipAnalysis.DEFAULT_SOURCE_FILENAME)
	# 		# print (pattern)
	# 		regex = re.compile(pattern)
	# 		for filename in os.listdir(obj.SourceFolderpath):
	# 			mobj = re.match(regex, filename)
	# 			if mobj is None: continue
	# 			# print (mobj.group(1))
	# 			filepath = os.path.join(obj.SourceFolderpath, filename)
	# 			if not os.path.isfile(filepath): continue
	# 			file_date = obj.get_data_date(mobj.group(1))
	# 			fliename_dict[file_date] = filename
	# 			# print(f)
	# 		# print (fliename_ordereddict)
	# 		fliename_ordereddict = OrderedDict(sorted(fliename_dict.items(), key=lambda x: x[0]))
	# 		for filedate, filename in fliename_ordereddict.items():
	# 			print ("%s: %s" % (filedate, filename))
	# 			# import pdb; pdb.set_trace()
	# 			obj.DatabaseDate = filedate
	# 			obj.SourceFilename = filename
	# 			obj.update_database()
	# 	elif args.find_database:
	# 		obj.find_database()
	# 	elif args.delete_database:
	# 		obj.delete_database()
	# 	elif args.list_database_date:
	# 		obj.list_database_date()
	# 	elif cfg['analysis_method'] == 0:
	# 		obj.search_sheets_from_file()
	# 	elif cfg['analysis_method'] == 1:
	# 		obj.search_sheets()
	# 	elif cfg['analysis_method'] == 2:
	# 		obj.search_sheets(True)
	# 	elif cfg['analysis_method'] == 3:
	# 		obj.evaluate_bull_bear()
	# 	else:
	# 		raise ValueError("Incorrect Analysis Method Index")
	cfg = {}
	with StockChipAnalysis(cfg) as obj:
		stock_chip_data_dict = obj.get_stock_chip_data()
		obj.search_targets(stock_chip_data_dict)
# 		stock_set1 = set(stock_chip_data_dict[u"主力買超天數累計"].keys())
# 		stock_set2 = set(stock_chip_data_dict[u"法人買超天數累計"].keys())
# 		stock_set3 = set(stock_chip_data_dict[u"外資買超天數累計"].keys())
# 		stock_set4 = set(stock_chip_data_dict[u"投信買超天數累計"].keys())
# # https://blog.csdn.net/qq_37195276/article/details/79467917
# # & != and ; | != or
# # python中&、|代表的是位運算符，and、or代表的是邏輯運算符
# 		stock_set = stock_set1 & stock_set2 & stock_set3 & stock_set4
# 		print (list(stock_set))
