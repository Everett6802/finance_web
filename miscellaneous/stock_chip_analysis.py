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
from datetime import datetime
from pymongo import MongoClient
from collections import OrderedDict


class StockChipAnalysis(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\\Users\\%s\\Downloads" % os.getlogin()
	DEFAULT_SOURCE_FILENAME = "stock_chip_analysis.xlsm"
	DEFAULT_CONFIG_FOLDERPATH =  "C:\\Users\\%s\\source" % os.getlogin()
	DEFAULT_STOCK_LIST_FILENAME = "chip_analysis_stock_list.txt"
	DEFAULT_REPORT_FILENAME = "chip_analysis_report.xlsx"
	DEFAULT_SEARCH_RESULT_FILENAME = "search_result_stock_list.txt"
	SHEET_METADATA_DICT = {
		# u"即時指數": { # Dummy
		# 	"is_dummy": True,
		# },
		# u"主要指數": { # Dummy
		# 	"is_dummy": True,
		# },
		# u"外匯市場": { # Dummy
		# 	"is_dummy": True,
		# },
		# u"商品市場": { # Dummy
		# 	"is_dummy": True,
		# },
		# u"商品行情": { # Dummy
		# 	"is_dummy": True,
		# },
		# u"資金流向": { # Dummy
		# 	"is_dummy": True,
		# },
		# u"大盤籌碼多空勢力": { # Dummy
		# 	"is_dummy": True,
		# },
		# u"焦點股": { 
		# 	"key_mode": 0, # 1476.TW
		# },
		u"法人共同買超累計": {
			"key_mode": 0, # 1476.TW
			"direction": "+",
		},
		u"主力買超天數累計": {
			"key_mode": 3, # 2504 國產
			"direction": "+",
		},
		u"法人買超天數累計": {
			"key_mode": 3, # 2504 國產
			"direction": "+",
		},
		u"外資買超天數累計": {
			"key_mode": 3, # 2504 國產
			"direction": "+",
		},
		u"投信買超天數累計": {
			"key_mode": 3, # 2504 國產
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
		u"買超異常": {
			"key_mode": 1, # 陽明(2609)
			"direction": "+",
		},
		u"賣超異常": {
			"key_mode": 1, # 陽明(2609)
			"direction": "-",
		},
		u"券商買最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "+",
		},
		u"券商賣最多股": {
			"key_mode": 1, # 陽明(2609)
			"direction": "-",
		},
		u"上市融資增加": {
			"key_mode": 2, # 4736  泰博
			"direction": "+",
		},
		u"上櫃融資增加": {
			"key_mode": 2, # 4736  泰博
			"direction": "-",
		},
	}
	ALL_SHEET_NAME_LIST = SHEET_METADATA_DICT.keys()
	DEFAULT_SHEET_NAME_LIST = [u"法人共同買超累計", u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計", u"外資買最多股", u"外資賣最多股", u"投信買最多股", u"投信賣最多股", u"主力買最多股", u"主力賣最多股", u"買超異常", u"賣超異常", u"券商買最多股", u"券商賣最多股", u"上市融資增加", u"上櫃融資增加",]
	SHEET_SET_LIST = [
		[u"法人共同買超累計", u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計",],
		[u"法人共同買超累計", u"外資買超天數累計", u"投信買超天數累計",],
		[u"外資買超天數累計", u"投信買超天數累計",],
	]
	DEFAULT_CONSECUTIVE_OVER_BUY_DAYS = 3
	CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET = [u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計",]
	CHECK_CONSECUTIVE_OVER_BUY_DAYS_FIELD_NAME_KEY = u"買超累計天數"

	WEIGHTED_STOCK_LIST = ["2330", "2317", "2454", "2308", "0050", ]
	DEFENSE_STOCK_LIST = ["2412", "3045", "4904", "2801", "2809", "2812", "2823", "2834", "2880", "2881", "2882", "2883", "2884", "2885", "2886", "2887", "2888", "2889", "2890", "2891", "2892", "5776", "5880", ]

	DEFAULT_DB_NAME = "StockChipAnalysis"
	DEFAULT_DB_USERNAME = "root"
	DEFAULT_DB_PASSWORD = "lab4man1"
	DEFAULT_DB_DATETIME_STRING_FORMAT = "%Y-%m-%d"

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


	# @classmethod
	# def read_stock_list_from_file(cls, stock_list_filepath):
	# 	# import pdb; pdb.set_trace()
	# 	if not cls.__check_file_exist(stock_list_filepath):
	# 		raise RuntimeError("The file[%s] does NOT exist" % stock_list_filepath)
	# 	stock_list = []
	# 	with open(stock_list_filepath, 'r') as fp:
	# 		for line in fp:
	# 			stock_list.append(line.strip("\n"))
	# 	return stock_list


	@classmethod
	def list_sheet_set(cls):
		for index, sheet_set in enumerate(cls.SHEET_SET_LIST):
			print ("%d: %s" % (index, ",".join(sheet_set)))


	def __init__(self, cfg):
		self.xcfg = {
			"show_detail": False,
			"generate_report": False,
			"source_filename": self.DEFAULT_SOURCE_FILENAME,
			"stock_list_filename": self.DEFAULT_STOCK_LIST_FILENAME,
			"report_filename": self.DEFAULT_REPORT_FILENAME,
			"stock_list": None,
			"sheet_name_list": None,
			"sheet_set_category": -1,
			"consecutive_over_buy_days": self.DEFAULT_CONSECUTIVE_OVER_BUY_DAYS,
			"need_all_sheet": False,
			"search_result_filename": self.DEFAULT_SEARCH_RESULT_FILENAME,
			"output_search_result": False,
			"quiet": False,
			"sort": False,
			"sort_limit": None,
			"db_host": "localhost",
			"db_name": self.DEFAULT_DB_NAME,
			"db_username": self.DEFAULT_DB_USERNAME,
			"db_password": self.DEFAULT_DB_PASSWORD,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["source_filepath"] = os.path.join(self.DEFAULT_SOURCE_FOLDERPATH, self.xcfg["source_filename"])
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

		self.workbook = None
		self.report_workbook = None
		self.search_result_txtfile = None
		self.sheet_title_bar_dict = {}
		self.db_client = None
		self.db_handle = None


	def __enter__(self):
		# Open the workbook
		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		if self.xcfg["generate_report"]:
			self.report_workbook = xlsxwriter.Workbook(self.xcfg["report_filepath"])
		if self.xcfg["output_search_result"]:
			self.search_result_txtfile = open(self.xcfg["search_result_filepath"], "w")

# mongodb://root:lab4man1@localhost:27017/StockChipAnalysis
		# db_url = 'mongodb://%s:%s@%s:27017' % (self.xcfg["db_username"], self.xcfg["db_password"], self.xcfg["db_host"])
		db_url = 'mongodb://%s:%s@%s:27017/%s' % (self.xcfg["db_username"], self.xcfg["db_password"], self.xcfg["db_host"], self.xcfg["db_name"])
		# print ("DB URL: %s" % db_url)
		self.db_client = MongoClient(db_url)
		# self.db_client = MongoClient('mongodb://%s:27017' % (self.xcfg["db_host"]))
# Database (Database -> Collection -> Document)
		self.db_handle = self.db_client[self.xcfg["db_name"]]
		# self.db_handle.authenticate(self.xcfg["db_username"], self.xcfg["db_password"])
		return self


	def __exit__(self, type, msg, traceback):
		if self.db_client is not None:
			self.db_handle = None
			self.db_client.close()
			self.db_client = None
		if self.search_result_txtfile is not None:
			self.search_result_txtfile.close()
			self.search_result_txtfile = None
		if self.report_workbook is not None:
			self.report_workbook.close()
			del self.report_workbook
			self.report_workbook = None
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


	def __print_string(self, outpug_str):
		if self.xcfg["quiet"]: return
		print (outpug_str)


	def __read_sheet_title_bar(self, sheet_name):
		# import pdb; pdb.set_trace()
# has_key has been deprecated in Python 3.0
		# if not self.sheet_title_bar_dict.has_key(sheet_name):
		if sheet_name not in self.sheet_title_bar_dict:
			sheet_metadata = self.SHEET_METADATA_DICT[sheet_name]
			worksheet = self.workbook.sheet_by_name(sheet_name)
			title_bar_list = [u"商品",]
			column_start_index = None
			if sheet_metadata["key_mode"] == 0:
				column_start_index = 2
			elif sheet_metadata["key_mode"] in [1, 2, 3,]:
				column_start_index = 1
			else:
				raise ValueError("Unknown key mode: %d" % sheet_metadata["key_mode"]) 
			for column_index in range(column_start_index, worksheet.ncols):
				title_bar_list.append(worksheet.cell_value(0, column_index))
			self.sheet_title_bar_dict[sheet_name] = title_bar_list
		return self.sheet_title_bar_dict[sheet_name]


	def __read_sheet_data(self, sheet_name, write_db=False):
		# import pdb; pdb.set_trace()
		sheet_metadata = self.SHEET_METADATA_DICT[sheet_name]
		# print (u"Read sheet: %s" % sheet_name)
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
			ignore_data = False
			stock_number = None
			# print "key_str: %s" % key_str
			if sheet_metadata["key_mode"] == 0:
				mobj = re.match("([\d]{4})\.TW", key_str)
				if mobj is None:
					raise ValueError("Incorrect format1: %s" % key_str)
				stock_number = mobj.group(1)
				data_dict[stock_number] = []
			elif sheet_metadata["key_mode"] == 1:
				mobj = re.match("(.+)\(([\d]{4}[\d]?[\w]?)\)", key_str)
				if mobj is None:
					raise ValueError("Incorrect format2: %s" % key_str)
				stock_number = mobj.group(2)
				data_dict[stock_number] = [mobj.group(1),]
			elif sheet_metadata["key_mode"] == 2:
				mobj = re.match("([\d]{4})\s{2}(.+)", key_str)
				if mobj is None:
					ignore_data = True
				else:
					stock_number = mobj.group(1)
					data_dict[stock_number] = [mobj.group(2),]
			elif sheet_metadata["key_mode"] == 3:
				mobj = re.match("([\d]{4})\s(.+)", key_str)
				if mobj is None:
					raise ValueError("Incorrect format3: %s" % key_str)
				stock_number = mobj.group(1)
				data_dict[stock_number] = [mobj.group(2),]
			else:
				raise ValueError("Unknown key mode: %d" % sheet_metadata["key_mode"])
			# if stock_number is None:
			#	raise RuntimeError("Fail to parse the stock number")
			if not ignore_data:
				for column_index in range(1, worksheet.ncols):
					data_dict[stock_number].append(worksheet.cell_value(row_index, column_index))
			row_index += 1
			# print "%d -- %s" % (row_index, stock_number)
		if self.xcfg["consecutive_over_buy_days"] > 0:
			if sheet_name in self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET:
				data_dict = self.__filter_by_consecutive_over_buy_days(sheet_name, data_dict)
		return data_dict


	def __filter_by_consecutive_over_buy_days(self, sheet_name, data_dict):
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
				# if not sheet_data_collection_dict.has_key(data_key):
				if data_key not in sheet_data_collection_dict:
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
		# import pdb; pdb.set_trace()
		for sheet_name in self.xcfg["sheet_name_list"]:
			data_dict = self.__read_sheet_data(sheet_name)
			for stock in self.xcfg["stock_list"]:
# has_key has been deprecated in Python 3.0
				# if not data_dict.has_key(stock):
				if stock not in data_dict:
					continue
				stock_data = data_dict[stock]
				# if not sheet_data_collection_dict.has_key(stock):
				if stock not in sheet_data_collection_dict:
					sheet_data_collection_dict[stock] = {}
				if sheet_data_func_ptr is not None:
					 stock_data = sheet_data_func_ptr(stock_data)
				sheet_data_collection_dict[stock][sheet_name] = stock_data
		return sheet_data_collection_dict


	def __filter_by_sheet_occurrence(self, sheet_data_collection_dict, sheet_occurrence_thres=None):
		assert self.xcfg["sheet_name_list"] is not None, "self.xcfg['sheet_name_list'] should NOT be None"
		# import pdb; pdb.set_trace()
		if sheet_occurrence_thres is None:
			sheet_occurrence_thres = len(self.xcfg["sheet_name_list"])
		# print "sheet_occurrence_thres: %d" % sheet_occurrence_thres
		# for stock_number, sheet_data_dict in sheet_data_collection_dict.items():
		# 	print "%s: %d" % (stock_number, len(sheet_data_dict.keys()))
		# import pdb; pdb.set_trace()
		return dict(filter(lambda x: len(x[1].keys()) == sheet_occurrence_thres, sheet_data_collection_dict.items()))


	def __sort_by_direction(self, sheet_data_collection_dict):
		stock_direction_statistics_list = []
		for stock_number, stock_sheet_data_collection_dict in sheet_data_collection_dict.items():
			# import pdb; pdb.set_trace()
			count = 0
			for sheet_name, _ in stock_sheet_data_collection_dict.items():
				# if not self.SHEET_METADATA_DICT[sheet_name].has_key("direction"):
				if "direction" not in self.SHEET_METADATA_DICT[sheet_name]:
					continue
				if self.SHEET_METADATA_DICT[sheet_name]["direction"] == "+":
					count += 1
				elif self.SHEET_METADATA_DICT[sheet_name]["direction"] == "-":
					count -= 1
			stock_direction_statistics_list.append((stock_number, count))
		stock_direction_statistics_list.sort(key=lambda x: x[1], reverse=True)
		# import pdb; pdb.set_trace()
		if self.xcfg["sort_limit"] is not None:
			stock_direction_statistics_list = stock_direction_statistics_list[:self.xcfg["sort_limit"]]
		sheet_data_collection_ordereddict = OrderedDict()
		for stock_number, _ in stock_direction_statistics_list:
			sheet_data_collection_ordereddict[stock_number] = sheet_data_collection_dict[stock_number]
		return sheet_data_collection_ordereddict


	def __search_stock_sheets(self):
		sheet_data_func_ptr = (lambda x: x) if self.xcfg["show_detail"] else (lambda x: x[0])
		sheet_data_collection_dict = self.__collect_sheet_data(sheet_data_func_ptr)
		if self.xcfg["need_all_sheet"]:
			sheet_data_collection_dict = self.__filter_by_sheet_occurrence(sheet_data_collection_dict)
		# import pdb; pdb.set_trace()
		if self.xcfg["sort"]:
			sheet_data_collection_dict = self.__sort_by_direction(sheet_data_collection_dict)
		if self.xcfg["stock_list"] is None:
			self.xcfg["stock_list"] = sheet_data_collection_dict.keys()
		else:
			if self.xcfg["sort"]:
				new_stock_list = filter(lambda x: x in self.xcfg["stock_list"], sheet_data_collection_dict.keys())
				self.xcfg["stock_list"] = new_stock_list

		no_data = True

		output_overview_worksheet = None
		output_overview_row = 0
		if self.xcfg["generate_report"]:
			output_overview_worksheet = self.report_workbook.add_worksheet("Overview")
					
		for stock_number in self.xcfg["stock_list"]:
			# if not sheet_data_collection_dict.has_key(stock_number):
			if stock_number not in sheet_data_collection_dict:
				continue
			no_data = False
			stock_sheet_data_collection_dict = sheet_data_collection_dict[stock_number]
			if self.xcfg["show_detail"]:
				stock_name = stock_sheet_data_collection_dict.values()[0][0]
				self.__print_string("=== %s(%s) ===" % (stock_number, stock_name))
				if self.xcfg["generate_report"]:
# For overview sheet
					output_overview_worksheet.write(output_overview_row, 0,  "%s(%s)" % (stock_number, stock_name))
					for output_overview_col, sheet_name in enumerate(stock_sheet_data_collection_dict.keys()):
						output_overview_worksheet.write(output_overview_row + 1, output_overview_col,  sheet_name)
					output_overview_row += 3
# For detailed sheet
					try:
						worksheet = self.report_workbook.add_worksheet("%s(%s)" % (stock_number, stock_name))
					except xlsxwriter.exceptions.InvalidWorksheetName:
						import pdb; pdb.set_trace()
						if re.match("6741", stock_number):
							worksheet = self.report_workbook.add_worksheet("%s(%s)" % (stock_number, stock_name.replace("*","")))
					output_row = 0
				for sheet_name, sheet_data_list in stock_sheet_data_collection_dict.items():
					sheet_title_bar_list = self.__read_sheet_title_bar(sheet_name)
					sheet_data_list_len = len(sheet_data_list)
					sheet_title_bar_list_len = len(sheet_title_bar_list)
					assert sheet_data_list_len == sheet_title_bar_list_len, "The list lengths are NOT identical, sheet_data_list_len: %d, sheet_title_bar_list_len: %d" % (sheet_data_list_len, sheet_title_bar_list_len)
					self.__print_string("* %s" % sheet_name)
					self.__print_string("%s" % ",".join(["%s[%s]" % elem for elem in zip(sheet_title_bar_list[1:], sheet_data_list[1:])]))
					if self.xcfg["generate_report"]:
# For detailed sheet
						worksheet.write(output_row, 0,  sheet_name)
						for output_col, output_data in enumerate(zip(sheet_title_bar_list[1:], sheet_data_list[1:])):
							sheet_title_bar, sheet_data = output_data
							worksheet.write(output_row + 1, output_col,  sheet_title_bar)
							worksheet.write(output_row + 2, output_col,  sheet_data)
						output_row += 4
			else:
				# import pdb; pdb.set_trace()
# For python 3, it's required to convert to list for the return value of the values function.
				stock_name = list(stock_sheet_data_collection_dict.values())[0]
				self.__print_string("=== %s(%s) ===" % (stock_number, stock_name))
				self.__print_string("%s" % (u",".join([stock_sheet_data_key for stock_sheet_data_key in stock_sheet_data_collection_dict.keys()])))
			if self.xcfg["output_search_result"]:
				self.search_result_txtfile.write("%s\n" % stock_number)
		if no_data: self.__print_string("*** No Data ***")	
		if self.xcfg["generate_report"]:
			if no_data:
				worksheet = self.report_workbook.add_worksheet("NoData")


	def __buy_sell_statistics(self, stock_list):
		buy_count = 0
		sell_count = 0
		sheet_name_list = dict(filter(lambda x: "direction" in x[1], self.SHEET_METADATA_DICT.items())).keys()
		# import pdb; pdb.set_trace()
		for sheet_name in sheet_name_list:
			data_dict = self.__read_sheet_data(sheet_name)
			count = len(filter(lambda x: x in stock_list, data_dict.keys()))
			if self.SHEET_METADATA_DICT[sheet_name]["direction"] == "+":
				buy_count += count
			else:
				sell_count += count
		return buy_count, sell_count


	def __get_db_date(self, db_date):
		db_date_obj = None
		if db_date is None: 
			now = datetime.now()
			db_date_obj = datetime(now.year, now.month, now.day)
		else:
			if type(db_date) is str:
				db_date_obj = datetime.strptime(str(db_date), self.DEFAULT_DB_DATETIME_STRING_FORMAT)
			else:
				raise ValueError("Unsupported type of the db_date object: %s" % type(db_date))
		return db_date_obj


	def __insert_db(self, sheet_data_collection_dict, db_date=None):
		assert self.db_handle is not None, "self.db_handle should NOT be None"
		# import pdb; pdb.set_trace()
		db_date = self.__get_db_date(db_date)
		insert_data_dict = {}
		for stock_number, stock_sheet_data_collection_dict in sheet_data_collection_dict.items():
			for sheet_name, stock_sheet_data_collection in stock_sheet_data_collection_dict.items():
				if sheet_name not in insert_data_dict:
					insert_data_dict[sheet_name] = {}
				insert_data_dict[sheet_name][stock_number] = stock_sheet_data_collection
		# import pdb; pdb.set_trace()
		for db_sheet_name, db_sheet_data_collection in insert_data_dict.items():
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			insert_data_dict = {
				"created_date": db_date,
				"data": db_sheet_data_collection,
			}
			# import pdb; pdb.set_trace()
# Insert Document
			# print ("================= %s =================" % db_sheet_name)
			# print (insert_data_dict)		
			db_collection_handle.insert_one(insert_data_dict)


	def __find_db(self, db_date=None):
		assert self.db_handle is not None, "self.db_handle should NOT be None"
		'''
		Data format:
		   {
		      sheet_name1: {
		         "company_no1": [value1, value2,...],
		         "company_no2": [value1, value2,...],
		         ...
		      },
		      sheet_name2: {
		         "company_no1": [value1, value2,...],
		         "company_no2": [value1, value2,...],
		         ...
		      },
		      ...
		   }
		'''
		# import pdb; pdb.set_trace()
		db_date = self.__get_db_date(db_date)
		find_criteria_dict = {
			"created_date": db_date,
		}
		find_data_dict = {}
		for db_sheet_name in self.ALL_SHEET_NAME_LIST:
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			#import pdb; pdb.set_trace()
# Find Document
			search_res = db_collection_handle.find(find_criteria_dict)
			search_res_cnt = search_res.count()
			if search_res_cnt > 1:
				raise ValueError("Incorrect data in %s: %d" % (db_sheet_name, search_res_cnt))
			# print ("================= %s ================= %d " % (db_sheet_name, search_res_cnt))
			if search_res_cnt == 0:
				continue
			for entry in search_res:
				# print (entry["data"])
				find_data_dict[db_sheet_name] = entry["data"]
		for sheet_name, data_dict in find_data_dict.items():
			print ("================= %s =================\n %s" % (sheet_name, data_dict))
		return find_data_dict


	def update_database(self):
		self.xcfg["consecutive_over_buy_days"] = 0
		sheet_data_collection_dict = self.__collect_sheet_all_data()
		self.__insert_db(sheet_data_collection_dict) 
		# find_data_dict = self.__find_db()
		# print (find_data_dict)


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


	def evaluate_bull_bear(self):
		buy_count, sell_count = self.__buy_sell_statistics(self.WEIGHTED_STOCK_LIST)
		print ("Weighted Stock, buy: %d, sell: %d" % (buy_count, sell_count))
		buy_count, sell_count = self.__buy_sell_statistics(self.DEFENSE_STOCK_LIST)
		print ("Defense Stock, buy: %d, sell: %d" % (buy_count, sell_count))


	@property
	def StockList(self):
		return self.xcfg["stock_list"]


	@StockList.setter
	def StockList(self, stock_list):
		self.xcfg["stock_list"] = stock_list


	@property
	def StockListFilepath(self):
		return self.xcfg["stock_list_filepath"]


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
	parser.add_argument('-i', '--list_sheet_set_category', required=False, action='store_true', help='List each stock set and exit')
	parser.add_argument('--update_database', required=False, action='store_true', help='Update database and exit')
	parser.add_argument('-p', '--create_report_by_sheet_set_category', required=False, help='Create a report by certain a sheet set category and exit')
	parser.add_argument('-m', '--analysis_method', required=False, help='The method for chip analysis. Default: 0')	
	parser.add_argument('-d', '--show_detail', required=False, action='store_true', help='Show detailed data for each stock')
	parser.add_argument('-g', '--generate_report', required=False, action='store_true', help='Generate the report of the detailed data for each stock to the XLS file.')
	parser.add_argument('-r', '--report_filename', required=False, help='The filename of chip analysis report')
	parser.add_argument('-t', '--stock_list_filename', required=False, help='The filename of stock list for chip analysis')
	parser.add_argument('-l', '--stock_list', required=False, help='The list string of stock list for chip analysis. Ex: 2330,2317,2454,2308')
	parser.add_argument('-u', '--source_filename', required=False, help='The filename of chip analysis data source')
	parser.add_argument('-c', '--sheet_set_category', required=False, help='The category for sheet set. Default: 0')	
	parser.add_argument('-n', '--need_all_sheet', required=False, action='store_true', help='The stock should be found in all sheets in the sheet name list')
	parser.add_argument('-a', '--search_result_filename', required=False, help='The filename of stock list for search result')
	parser.add_argument('-o', '--output_search_result', required=False, action='store_true', help='Ouput the search result')
	parser.add_argument('-q', '--quiet', required=False, action='store_true', help="Don't print string on the screen")
	parser.add_argument('-s', '--sort', required=False, action='store_true', help="Show the data in order")
	parser.add_argument('-f', '--sort_limit', required=False, help="Limit the sorted data")
	# parser.add_argument('-f', '--update_datebase', required=False, help="Limit the sorted data")
	args = parser.parse_args()

	if args.list_analysis_method:
		help_str_list = [
			"Search sheet for the specific stocks from the file",
			"Search sheet for the specific stocks",
			"Search sheet for the whole stocks",
			"Evaluate the TAIEX bull or bear",
		]
		help_str_list_len = len(help_str_list)
		print ("************ Analysis Method ************")
		for index, help_str in enumerate(help_str_list):
			print ("%d  %s" % (index, help_str))
		print ("*****************************************")
		sys.exit(0)
	if args.list_sheet_set_category:
		StockChipAnalysis.list_sheet_set()
		sys.exit(0)
	if args.update_database:
		with StockChipAnalysis({}) as obj: 
			obj.update_database()
		sys.exit(0)
	if args.create_report_by_sheet_set_category:
		search_result_filename = "tmp1.txt"
		cfg_step1 = {
			"sheet_set_category": int(args.create_report_by_sheet_set_category),
			"need_all_sheet": True,
			"search_result_filename": search_result_filename,
			"output_search_result": True,
			"quiet": True,
		}
		with StockChipAnalysis(cfg_step1) as obj_step1: 
			obj_step1.search_sheets(True)
		cfg_step2 = {
			"generate_report": True,
			"stock_list_filename": search_result_filename,
			"quiet": True,
		}
		with StockChipAnalysis(cfg_step2) as obj_step2: 
			obj_step2.search_sheets_from_file()
			os.remove(obj_step2.StockListFilepath)
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
	cfg['sheet_set_category'] = int(args.sheet_set_category) if args.sheet_set_category is not None else -1
	if args.need_all_sheet: cfg['need_all_sheet'] = True
	if args.report_filename is not None: cfg['report_filename'] = args.report_filename
	if args.search_result_filename is not None: cfg['search_result_filename'] = args.search_result_filename
	if args.output_search_result: cfg['output_search_result'] = True
	if args.quiet: cfg['quiet'] = True
	if args.sort: cfg['sort'] = True
	if args.sort_limit is not None: cfg['sort_limit'] = int(args.sort_limit)
		
	# import pdb; pdb.set_trace()
	with StockChipAnalysis(cfg) as obj:
		if cfg['analysis_method'] == 0:
			obj.search_sheets_from_file()
		elif cfg['analysis_method'] == 1:
			obj.search_sheets()
		elif cfg['analysis_method'] == 2:
			obj.search_sheets(True)
		elif cfg['analysis_method'] == 3:
			obj.evaluate_bull_bear()
		else:
			raise ValueError("Incorrect Analysis Method Index")
