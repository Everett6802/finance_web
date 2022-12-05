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
	DEFAULT_SOURCE_FILENAME = "stock_chip_analysis"
	DEFAULT_SOURCE_FULL_FILENAME = "%s.xlsm" % DEFAULT_SOURCE_FILENAME
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
	DEFAULT_CONSECUTIVE_OVER_BUY_DAYS = 0
	CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET = [u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計",]
	CHECK_CONSECUTIVE_OVER_BUY_DAYS_FIELD_NAME_KEY = u"買超累計天數"

	WEIGHTED_STOCK_LIST = ["2330", "2317", "2454", "2308", "0050", ]
	DEFENSE_STOCK_LIST = ["2412", "3045", "4904", "2801", "2809", "2812", "2823", "2834", "2880", "2881", "2882", "2883", "2884", "2885", "2886", "2887", "2888", "2889", "2890", "2891", "2892", "5776", "5880", ]

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


	@classmethod
	def get_data_date(cls, data_date=None):
		data_date_obj = None
		if data_date is None: 
			now = datetime.now()
			data_date_obj = datetime(now.year, now.month, now.day)
		else:
			if type(data_date) is str:
				data_date_obj = datetime.strptime(str(data_date), cls.DEFAULT_DB_DATE_STRING_FORMAT)
			elif type(data_date) is datetime:
				data_date_obj = data_date
			else:
				raise ValueError("Unsupported type of the data_date object: %s" % type(data_date))
		return data_date_obj


	@classmethod
	def __read_from_worksheet(cls, worksheet, sheet_metadata):
		data_dict = {}
		row_index = 1
		while True:
			try:
				key_str = worksheet.cell_value(row_index, 0)
			except IndexError as e:
				# print ("Fail to read data [%s] in (%d, 0), due to: %s" % (sheet_name, row_index, str(e)))
				break
			ignore_data = False
			stock_number = None
			# print "key_str: %s" % key_str
			if sheet_metadata["key_mode"] == 0:
				# mobj = re.match("([\d]{4})\.TW", key_str)
				mobj = re.match("([\d]{4})", str(int(key_str)))
				if mobj is None:
					raise ValueError("%s: Incorrect format1: %s" % (sheet_name, key_str))
				stock_number = mobj.group(1)
				data_dict[stock_number] = []
			elif sheet_metadata["key_mode"] == 1:
				# mobj = re.match("(.+)\(([\d]{4}[\d]?[\w]?)\)", key_str)
				mobj = re.match("(.+)\(([\d]{4}[\d\w]{0,2}.*)\)", key_str)
				# print ("%s: %s" % (sheet_name, key_str))
				if mobj is None:
					import pdb; pdb.set_trace()
					raise ValueError("%s: Incorrect format2: %s" % (sheet_name, key_str))
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
					raise ValueError("%s: Incorrect format3: %s" % (sheet_name, key_str))
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
		return data_dict


	@classmethod
	def check_data_source_difference(cls, data_date_str=None):
		# import pdb; pdb.set_trace()
		data_date = cls.get_data_date(data_date_str)
# data from XLS
		source_filename = "%s@%s.xlsm" % (cls.DEFAULT_SOURCE_FILENAME, data_date.strftime(cls.DEFAULT_DB_DATE_STRING_FORMAT))
		xls_filepath = os.path.join(cls.DEFAULT_SOURCE_FOLDERPATH, source_filename)
		xls_workbook = xlrd.open_workbook(xls_filepath)
# data from DB
# mongodb://root:lab4man1@localhost:27017/StockChipAnalysis
		db_url = 'mongodb://%s:%s@%s:27017/%s' % (cls.DEFAULT_DB_USERNAME, cls.DEFAULT_DB_PASSWORD, "localhost", cls.DEFAULT_DB_NAME)
		db_client = MongoClient(db_url)
		db_handle = db_client[cls.DEFAULT_DB_NAME]
		db_criteria_dict = {
			"created_date": data_date,
		}

		print (data_date.strftime(cls.DEFAULT_DB_DATE_STRING_FORMAT))
		for sheet_name in cls.ALL_SHEET_NAME_LIST:
			# import pdb; pdb.set_trace()
			xls_worksheet = xls_workbook.sheet_by_name(sheet_name)
			sheet_metadata = cls.SHEET_METADATA_DICT[sheet_name]
			xls_data_dict = cls.__read_from_worksheet(xls_worksheet, sheet_metadata)
			db_collection_handle = db_handle[sheet_name]
			search_res = db_collection_handle.find(db_criteria_dict)
			db_data_dict = search_res[0]["data"]
			print (("=" * 10 + " %s " + "=" * 10) % sheet_name)
			# print ("* XLS")
			# print (xls_data_dict)
			# print ("* DB")
			# print (db_data_dict)
			print ("Equal" if xls_data_dict == db_data_dict else "Not Equal")


		if consecutive_over_buy_days is None:
			consecutive_over_buy_days = self.xcfg["consecutive_over_buy_days"]
		if consecutive_over_buy_days > 0:
			if sheet_name in self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET:
				data_dict = self.__filter_by_consecutive_over_buy_days(data_dict, sheet_name=sheet_name)
		return data_dict


		find_data_dict = {}
		for db_sheet_name in self.ALL_SHEET_NAME_LIST:
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			# import pdb; pdb.set_trace()
# Find Document
			search_res = db_collection_handle.find(find_criteria_dict)
# Deprecated. The behaviour differed (estimated vs actual count) based on whether query criteria was provided
			# search_res_cnt = search_res.count()  
			search_res_cnt = db_collection_handle.count_documents(find_criteria_dict)
			# if search_res_cnt > 1:
			# 	raise ValueError("Incorrect data in %s: %d" % (db_sheet_name, search_res_cnt))
			# print ("================= %s ================= %d " % (db_sheet_name, search_res_cnt))
			stock_list_not_empty = (self.xcfg["stock_list"] is not None)
			if search_res_cnt != 0:
				if ret_date_first:
					for entry in search_res:
						'''
						entry --
						       |-'_id'
						       |-'created_date'
						       |-'data'
						       |-'metadata'
						'''
						# print (entry["created_date"])
						# print (entry["data"])
						if entry["created_date"] not in find_data_dict:
							find_data_dict[entry["created_date"]] = {} 
						if data_for_analysis:
							'''
							find_data_dict -- created_date
											       |-'data'
											       |-'metadata'
							'''
							find_data_dict[entry["created_date"]]["data"] = {}
							find_data_dict[entry["created_date"]]["metadata"] = entry["metadata"]
							if self.xcfg["consecutive_over_buy_days"] > 0:
								if db_sheet_name in self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET:
									entry["data"] = self.__filter_by_consecutive_over_buy_days(entry["data"], title_bar_list=entry["metadata"][db_sheet_name])
							for stock_number, stock_data in entry["data"].items():
								if stock_list_not_empty and stock_number not in self.xcfg["stock_list"]:
									continue
								if stock_number not in find_data_dict[entry["created_date"]]["data"]:
									find_data_dict[entry["created_date"]]["data"][stock_number] = {}
								if sheet_data_func_ptr is not None:
									 stock_data = sheet_data_func_ptr(stock_data)
								find_data_dict[entry["created_date"]]["data"][stock_number][db_sheet_name] = stock_data
						else:
							find_data_dict[entry["created_date"]][db_sheet_name] = entry["data"]



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
			"consecutive_over_buy_days": self.DEFAULT_CONSECUTIVE_OVER_BUY_DAYS,
			"need_all_sheet": False,
			"search_history": False,
			"search_result_filename": self.DEFAULT_SEARCH_RESULT_FILENAME,
			"output_search_result": False,
			"quiet": False,
			"sort": False,
			"sort_limit": None,
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


	def __read_sheet_title_bar(self, sheet_name):
		# import pdb; pdb.set_trace()
# has_key has been deprecated in Python 3.0
		# if not self.sheet_title_bar_dict.has_key(sheet_name):
		if sheet_name not in self.sheet_title_bar_dict:
			sheet_metadata = self.SHEET_METADATA_DICT[sheet_name]
			worksheet = self.__get_workbook().sheet_by_name(sheet_name)
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


	def __read_sheet_data(self, sheet_name, consecutive_over_buy_days=None):
		# import pdb; pdb.set_trace()
		sheet_metadata = self.SHEET_METADATA_DICT[sheet_name]
		# print (u"Read sheet: %s" % sheet_name)
		# assert self.workbook is not None, "self.workbook should NOT be None"
		worksheet = self.__get_workbook().sheet_by_name(sheet_name)
		# https://www.itread01.com/content/1549650266.html
		# print worksheet.name,worksheet.nrows,worksheet.ncols    #Sheet1 6 4
# The data
		data_dict = self.__read_from_worksheet(worksheet, sheet_metadata)

		if consecutive_over_buy_days is None:
			consecutive_over_buy_days = self.xcfg["consecutive_over_buy_days"]
		if consecutive_over_buy_days > 0:
			if sheet_name in self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET:
				data_dict = self.__filter_by_consecutive_over_buy_days(data_dict, sheet_name=sheet_name)
		return data_dict


	def __filter_by_consecutive_over_buy_days(self, data_dict, sheet_name=None, title_bar_list=None):
		# import pdb; pdb.set_trace()
		if title_bar_list is None:
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


# 	def __collect_sheet_all_data(self, sheet_data_func_ptr=None):
# 		sheet_data_collection_dict = {}
# 		if self.xcfg["sheet_name_list"] is None:
# 			self.xcfg["sheet_name_list"] = self.DEFAULT_SHEET_NAME_LIST
# 		sheet_data_collection_dict["metadata"] = {}
# 		sheet_data_collection_dict["data"] = {}
# 		for sheet_name in self.xcfg["sheet_name_list"]:
# 			# print (sheet_name)
# # metadata
# 			sheet_title_bar_list = self.__read_sheet_title_bar(sheet_name)
# 			sheet_data_collection_dict["metadata"][sheet_name] = sheet_title_bar_list
# # data
# 			data_dict = self.__read_sheet_data(sheet_name)
# 			for data_key, data_value in data_dict.items():
# 				# if not sheet_data_collection_dict.has_key(data_key):
# 				if data_key not in sheet_data_collection_dict["data"]:
# 					sheet_data_collection_dict["data"][data_key] = {}
# 				if sheet_data_func_ptr is not None:
# 					 data_value = sheet_data_func_ptr(data_value)
# 				sheet_data_collection_dict["data"][data_key][sheet_name] = data_value
# 		return sheet_data_collection_dict


	def __collect_sheet_data(self, data_for_analysis=False, sheet_data_func_ptr=None):
		# if self.xcfg["stock_list"] is None:
		# 	return self.__collect_sheet_all_data(sheet_data_func_ptr)
		sheet_data_collection_dict = {}
		if self.xcfg["sheet_name_list"] is None:
			self.xcfg["sheet_name_list"] = self.DEFAULT_SHEET_NAME_LIST
		# import pdb; pdb.set_trace()
		sheet_data_collection_dict["metadata"] = {}
		sheet_data_collection_dict["data"] = {}
		for sheet_name in self.xcfg["sheet_name_list"]:
# metadata
			sheet_title_bar_list = self.__read_sheet_title_bar(sheet_name)
			sheet_data_collection_dict["metadata"][sheet_name] = sheet_title_bar_list
# data
			data_dict = self.__read_sheet_data(sheet_name)
			stock_list = None
			if self.xcfg["stock_list"] is None:
				stock_list = data_dict.keys()
			else:
# has_key has been deprecated in Python 3.0
				# if not data_dict.has_key(stock):
				# stock_list = [stock for stock in self.xcfg["stock_list"] if stock in data_dict]
				stock_list = list(set(self.xcfg["stock_list"]) | set(data_dict.keys()))
			if data_for_analysis:
# Find the stock list
				for stock_number in stock_list:
					# stock_data = data_dict[stock]
					# if stock not in sheet_data_collection_dict["data"]:
					# 	sheet_data_collection_dict["data"][stock] = {}
					# # if sheet_data_func_ptr is not None:
					# # 	 stock_data = sheet_data_func_ptr(stock_data)
					# sheet_data_collection_dict["data"][stock][sheet_name] = stock_data
					sheet_data_collection_dict["data"].setdefault(stock_number, {}).update({sheet_name: data_dict[stock_number]})
			else:
				sheet_data_collection_dict["data"][sheet_name] = dict(filter(lambda x: x[0] in stock_list, data_dict.items()))
		# import pdb; pdb.set_trace()
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


	def check_data_source_difference(self, data_for_analysis=False):
		# import pdb; pdb.set_trace()
		sheet_data_func_ptr = (lambda x: x) if self.xcfg["show_detail"] else (lambda x: x[0])
		db_sheet_data_collection_dict_history = self.__find_db(self.xcfg["database_date"], data_for_analysis=data_for_analysis, sheet_data_func_ptr=sheet_data_func_ptr)
		# import pdb; pdb.set_trace()
		xls_sheet_data_collection_dict = self.__collect_sheet_data(data_for_analysis=data_for_analysis, sheet_data_func_ptr=sheet_data_func_ptr)
		xls_sheet_data_collection_dict_history = {self.get_data_date(self.xcfg["database_date"]): xls_sheet_data_collection_dict}
		# print (db_sheet_data_collection_dict_history.values())
		if data_for_analysis:
			db_data_dict = {}
			for item in db_sheet_data_collection_dict_history.values():
				for stock, data in item["data"].items():
					# print ("************ %s ************" % sheet_name)
					# print (data)
					db_data_dict[stock] = data
			# print (xls_sheet_data_collection_dict_history.values())
			print ("\n$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$\n$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$\n")
			xls_data_dict = {}
			for item in xls_sheet_data_collection_dict_history.values():
				for stock, data in item["data"].items():
					# print ("************ %s ************" % sheet_name)
					# print (data)
					xls_data_dict[stock] = data
			# import pdb; pdb.set_trace()
			assert set(db_data_dict.keys()) == set(xls_data_dict.keys()), "The stock lists are NOT identical"
			stock_list = db_data_dict.keys()
			for stock in stock_list:
				print("================== %s ==================" % stock)
				print("Equal" if (db_data_dict[stock] == xls_data_dict[stock]) else "Not Equal")
		else:
			# import pdb; pdb.set_trace()
			db_data_dict = {}
			for item in db_sheet_data_collection_dict_history.values():
				for sheet_name, data in item["data"].items():
					# print ("************ %s ************" % sheet_name)
					# print (data)
					db_data_dict[sheet_name] = data
			# print (xls_sheet_data_collection_dict_history.values())
			print ("\n$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$\n$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$\n")
			# import pdb; pdb.set_trace()
			xls_data_dict = {}
			for item in xls_sheet_data_collection_dict_history.values():
				for sheet_name, data in item["data"].items():
					print ("************ %s ************" % sheet_name)
					# print (data)
					xls_data_dict[sheet_name] = data
			# import pdb; pdb.set_trace()
			for sheet_name in self.ALL_SHEET_NAME_LIST:
				print("================== %s ==================" % sheet_name)
				print("Equal" if (db_data_dict[sheet_name] == xls_data_dict[sheet_name]) else "Not Equal")


	def __search_stock_sheets(self):
		# import pdb; pdb.set_trace()
		sheet_data_func_ptr = (lambda x: x) if self.xcfg["show_detail"] else (lambda x: x[0])
		sheet_data_collection_dict_history = None
		if self.xcfg["search_history"]:
			if self.xcfg["database_date"] is not None:
				sheet_data_collection_dict_history = self.__find_db(self.xcfg["database_date"], ret_date_first=True, data_for_analysis=True, sheet_data_func_ptr=sheet_data_func_ptr)
			elif self.xcfg["database_all_date_range"]:
				sheet_data_collection_dict_history = self.__find_db_range(ret_date_first=True, data_for_analysis=True, sheet_data_func_ptr=sheet_data_func_ptr)
			elif self.xcfg["database_date_range"] is not None:
				sheet_data_collection_dict_history = self.__find_db_range(self.xcfg["database_date_range_start"], self.xcfg["database_date_range_end"], ret_date_first=True, data_for_analysis=True, sheet_data_func_ptr=sheet_data_func_ptr)
			else:
				raise ValueError("Should select a date if the data source is from the databases")
		else:
# The data read from XLS is different from the one from DB 
# if the consecutive_over_buy_day is NOT 0 (default: DEFAULT_CONSECUTIVE_OVER_BUY_DAYS)
			sheet_data_collection_dict = self.__collect_sheet_data(sheet_data_func_ptr)
			data_date = self.get_data_date()
			sheet_data_collection_dict_history = {data_date: sheet_data_collection_dict}
			'''
			sheet_data_collection_dict_history -- created_date
												       |-'data'
												       |-'metadata'
			'''
		print (sheet_data_collection_dict_history.values())
		import pdb; pdb.set_trace()
		for data_date, sheet_data_collection_dict in sheet_data_collection_dict_history.items():
			if self.xcfg["need_all_sheet"]:
				sheet_data_collection_dict = self.__filter_by_sheet_occurrence(sheet_data_collection_dict["data"])
			# import pdb; pdb.set_trace()
			if self.xcfg["sort"]:
				sheet_data_collection_dict["data"] = self.__sort_by_direction(sheet_data_collection_dict["data"])
			stock_list = self.xcfg["stock_list"]
			if stock_list is None:
				stock_list = sheet_data_collection_dict.keys()
			else:
				if self.xcfg["sort"]:
					new_stock_list = filter(lambda x: x in stock_list, sheet_data_collection_dict.keys())
					stock_list = new_stock_list
			sheet_data_collection_dict_history[data_date] = sheet_data_collection_dict

			report_workbook = None
			if self.xcfg["generate_report"]:
				xls_filename = "%s-%s" % (self.xcfg["report_filepath"], data_date.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT))
				report_workbook = xlsxwriter.Workbook(xls_filename)

			no_data = True
			output_overview_worksheet = None
			output_overview_row = 0
			if self.xcfg["generate_report"]:
				output_overview_worksheet = report_workbook.add_worksheet("Overview")
# Output the data
			self.__print_string("********** %s **********" % data_date.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT))
			if self.xcfg["output_search_result"]:
				self.search_result_txtfile.write("********** %s **********\n" % data_date.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT))
			for stock_number in self.xcfg["stock_list"]:
				# if not sheet_data_collection_dict.has_key(stock_number):
				if stock_number not in sheet_data_collection_dict:
					continue
				no_data = False
				stock_sheet_data_collection_dict = sheet_data_collection_dict[stock_number]
				# import pdb; pdb.set_trace()
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
							worksheet = report_workbook.add_worksheet("%s(%s)" % (stock_number, stock_name))
						except xlsxwriter.exceptions.InvalidWorksheetName:
							# import pdb; pdb.set_trace()
							if re.match("6741", stock_number):
								worksheet = report_workbook.add_worksheet("%s(%s)" % (stock_number, stock_name.replace("*","")))
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
					worksheet = report_workbook.add_worksheet("NoData")
			if report_workbook is not None:
				report_workbook.close()
				del report_workbook
				report_workbook = None


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


	def __insert_db(self, sheet_data_collection_dict, data_date=None):
		assert self.db_handle is not None, "self.db_handle should NOT be None"
		# import pdb; pdb.set_trace()
		data_date = self.get_data_date(data_date)
		insert_data_dict = {}
		for stock_number, stock_sheet_data_collection_dict in sheet_data_collection_dict["data"].items():
			for sheet_name, stock_sheet_data_collection in stock_sheet_data_collection_dict.items():
				if sheet_name not in insert_data_dict:
					insert_data_dict[sheet_name] = {}
				insert_data_dict[sheet_name][stock_number] = stock_sheet_data_collection
		# import pdb; pdb.set_trace()
		for db_sheet_name, db_sheet_data_collection in insert_data_dict.items():
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			insert_data_dict = {
				"created_date": data_date,
				"data": db_sheet_data_collection,
				"metadata": sheet_data_collection_dict["metadata"],
			}
			# import pdb; pdb.set_trace()
# Insert Document
			# print ("================= %s =================" % db_sheet_name)
			# print (insert_data_dict)		
			ret = db_collection_handle.insert_one(insert_data_dict)
			# print (ret)

# data_for_analysis/sheet_data_func_ptr only takes effect when ret_date_first is True
# The metadata is only included when data_for_analysis is true
	def __find_db_internal(self, find_criteria_dict, check_exist_only=False, ret_date_first=True, data_for_analysis=False, sheet_data_func_ptr=None):
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
		find_data_dict = {}
		for db_sheet_name in self.ALL_SHEET_NAME_LIST:
			# print("DB sheet name: %s" % db_sheet_name)
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			# import pdb; pdb.set_trace()
# Find Document
			search_res = db_collection_handle.find(find_criteria_dict)
# Deprecated. The behaviour differed (estimated vs actual count) based on whether query criteria was provided
			# search_res_cnt = search_res.count()  
			search_res_cnt = db_collection_handle.count_documents(find_criteria_dict)
			# if search_res_cnt > 1:
			# 	raise ValueError("Incorrect data in %s: %d" % (db_sheet_name, search_res_cnt))
			# print ("================= %s ================= %d " % (db_sheet_name, search_res_cnt))
			stock_list_not_empty = (self.xcfg["stock_list"] is not None)
			if search_res_cnt != 0:
				if ret_date_first:
					for entry in search_res:
						'''
						entry --
						       |-'_id'
						       |-'created_date'
						       |-'data'
						       |-'metadata'
						'''
						# print (entry["created_date"])
						# print (entry["data"])
						if entry["created_date"] not in find_data_dict:
							find_data_dict[entry["created_date"]] = {} 
							find_data_dict[entry["created_date"]]["metadata"] = entry["metadata"]
							find_data_dict[entry["created_date"]]["data"] = {}
						'''
						find_data_dict -- created_date
											       |-'metadata'
											       |-'data'
						'''
						# find_data_dict[entry["created_date"]]["metadata"][db_sheet_name] = entry["metadata"]
						# if self.xcfg["consecutive_over_buy_days"] > 0:
						# 	if db_sheet_name in self.CHECK_CONSECUTIVE_OVER_BUY_DAYS_SHEET_SET:
						# 		entry["data"] = self.__filter_by_consecutive_over_buy_days(entry["data"], title_bar_list=entry["metadata"][db_sheet_name])
						for stock_number, stock_data in entry["data"].items():
							# if stock_list_not_empty and stock_number not in self.xcfg["stock_list"]:
							# 	continue
							if data_for_analysis:							
								# if stock_number not in find_data_dict[entry["created_date"]]["data"]:
								# 	find_data_dict[entry["created_date"]]["data"][stock_number] = {}
								# if sheet_data_func_ptr is not None:
								# 	 stock_data = sheet_data_func_ptr(stock_data)
								# find_data_dict[entry["created_date"]]["data"][stock_number][db_sheet_name] = stock_data
								find_data_dict[entry["created_date"]]["data"].setdefault(stock_number, {}).update({db_sheet_name: stock_data})
							else:
								# if sheet_data_func_ptr is not None:
								# 	 stock_data = sheet_data_func_ptr(stock_data)
								# find_data_dict[entry["created_date"]]["data"][db_sheet_name][stock_number] = stock_data
								find_data_dict[entry["created_date"]]["data"].setdefault(db_sheet_name, {}).update({stock_number: stock_data})
						# import pdb; pdb.set_trace()
						# print(find_data_dict[entry["created_date"]]["data"].keys())
				else:
					find_data_dict[db_sheet_name] = {}
					for entry in search_res:
						# print (entry["created_date"])
						# print (entry["data"])
						find_data_dict[db_sheet_name][entry["created_date"]] = entry["data"]
				if check_exist_only:
					break
		# for sheet_name, data_dict in find_data_dict.items():
		# 	print ("================= %s =================\n %s" % (sheet_name, data_dict))
		# import pdb; pdb.set_trace()
		return find_data_dict


	def __find_db(self, data_date=None, check_exist_only=False, ret_date_first=True, data_for_analysis=False, sheet_data_func_ptr=None):
		# import pdb; pdb.set_trace()
		data_date = self.get_data_date(data_date)
		find_criteria_dict = {
			"created_date": data_date,
		}
		return self.__find_db_internal(find_criteria_dict, check_exist_only, ret_date_first, data_for_analysis, sheet_data_func_ptr)


	def __find_db_range(self, start_data_date=None, end_data_date=None, check_exist_only=False, ret_date_first=True, data_for_analysis=False, sheet_data_func_ptr=None):
		find_criteria_dict = None
		if start_data_date is not None and end_data_date is not None:
			start_data_date = self.get_data_date(start_data_date)
			end_data_date = self.get_data_date(end_data_date)
			find_criteria_dict = {
				"$and":[{"created_date": {"$gte": start_data_date}}, {"created_date": {"$lte": end_data_date}}]
			}
		elif start_data_date is not None:
			start_data_date = self.get_data_date(start_data_date)
			find_criteria_dict = {
				"created_date": {"$gte": start_data_date}
			}
		elif end_data_date is not None:
			end_data_date = self.get_data_date(end_data_date)
			find_criteria_dict = {
				"created_date": {"$lte": end_data_date}
			}
		else:
			find_criteria_dict = {}
		return self.__find_db_internal(find_criteria_dict, check_exist_only, ret_date_first, data_for_analysis, sheet_data_func_ptr)


	def __check_db_data_exist(self, data_date=None):
		# import pdb; pdb.set_trace()
		find_data_dict = self.__find_db(data_date, check_exist_only=True)
# Empty dictionaries evaluate to False in Python:
		return bool(find_data_dict)


	def __update_db(self, sheet_data_collection_dict, data_date=None):
		assert self.db_handle is not None, "self.db_handle should NOT be None"
		# import pdb; pdb.set_trace()
		data_date = self.get_data_date(data_date)
		update_criteria_dict = {
			"created_date": data_date,
		}
		update_data_dict = {}
		for stock_number, stock_sheet_data_collection_dict in sheet_data_collection_dict["data"].items():
			for sheet_name, stock_sheet_data_collection in stock_sheet_data_collection_dict.items():
				if sheet_name not in update_data_dict:
					update_data_dict[sheet_name] = {}
				update_data_dict[sheet_name][stock_number] = stock_sheet_data_collection
		# import pdb; pdb.set_trace()
		for db_sheet_name, db_sheet_data_collection in update_data_dict.items():
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			update_data_dict = {
				"$set": {
					"data": db_sheet_data_collection,
					"metadata": sheet_data_collection_dict["metadata"],
				}
			}
			# import pdb; pdb.set_trace()
# Insert Document
			# print ("================= %s =================" % db_sheet_name)
			# print (insert_data_dict)		
			db_collection_handle.update_one(update_criteria_dict, update_data_dict)


	def __delete_db(self, data_date=None):
		assert self.db_handle is not None, "self.db_handle should NOT be None"
		# import pdb; pdb.set_trace()
		data_date = self.get_data_date(data_date)
		delete_criteria_dict = {
			"created_date": data_date,
		}
		for db_sheet_name in self.ALL_SHEET_NAME_LIST:
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			# import pdb; pdb.set_trace()
# Find Document
			delete_res = db_collection_handle.delete_one(delete_criteria_dict)
			# print ("================= %s =================\n %d Deleted" % (db_sheet_name, delete_res.deleted_count))


	def __delete_db_range(self, start_data_date=None, end_data_date=None):
		delete_criteria_dict = None
		if start_data_date is not None and end_data_date is not None:
			start_data_date = self.get_data_date(start_data_date)
			end_data_date = self.get_data_date(end_data_date)
			delete_criteria_dict = {
				"$and":[{"created_date": {"$gte": start_data_date}}, {"created_date": {"$lte": end_data_date}}]
			}
		elif start_data_date is not None:
			start_data_date = self.get_data_date(start_data_date)
			delete_criteria_dict = {
				"created_date": {"$gte": start_data_date}
			}
		elif end_data_date is not None:
			end_data_date = self.get_data_date(end_data_date)
			delete_criteria_dict = {
				"created_date": {"$lte": end_data_date}
			}
		else:
			delete_criteria_dict = {}
		for db_sheet_name in self.ALL_SHEET_NAME_LIST:
# Collection			
			db_collection_handle = self.db_handle[db_sheet_name]
			# import pdb; pdb.set_trace()
# Find Document
			delete_res = db_collection_handle.delete_many(delete_criteria_dict)


	def update_database(self):
		self.xcfg["consecutive_over_buy_days"] = 0
		# import pdb; pdb.set_trace()
		sheet_data_collection_dict = self.__collect_sheet_all_data()
		data_date = self.get_data_date(self.xcfg["database_date"])
		if self.__check_db_data_exist(data_date):
			print ("Data on %s already exists. Update the database..." % data_date.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT))
			self.__update_db(sheet_data_collection_dict, data_date)
		else:
			print ("Insert Data on %s to the database..." % data_date.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT))
			self.__insert_db(sheet_data_collection_dict, data_date) 


	def find_database(self, sheet_name_filter=[u"主力買超天數累計", u"法人買超天數累計"]):
		find_res = None
		if self.xcfg["database_date"] is not None:
			find_res = self.__find_db(self.xcfg["database_date"], ret_date_first=True)
		elif self.xcfg["database_all_date_range"]:
			find_res = self.__find_db_range(ret_date_first=True)
		elif self.xcfg["database_date_range"] is not None:
			find_res = self.__find_db_range(self.xcfg["database_date_range_start"], self.xcfg["database_date_range_end"], ret_date_first=True)
			# ret = self.__find_db_range("2022-06-01", "2022-06-02", ret_date_first=True)
		else:
			find_res = self.__find_db(None, ret_date_first=True)
		import pdb; pdb.set_trace()
		for databse_date, data_dict in find_res.items():
			print ("=============== %s ===============" % databse_date.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT))
			for sheet_name, sub_data_dict in data_dict.items():
				if (sheet_name_filter is not None) and (sheet_name not in sheet_name_filter):
					continue
				print ("***** %s *****" % sheet_name)
				print (sub_data_dict)


	def list_database_date(self, start_data_date=None, end_data_date=None, check_exist_only=False, ret_date_first=False, need_sort=True):
		find_data_dict = self.__find_db_range(start_data_date, end_data_date, check_exist_only=False, ret_date_first=False)
		for collection_name, data_dict in find_data_dict.items():
			data_count = len(data_dict.keys())
			print ("========== %s ========== %d" % (collection_name, data_count))
			if data_count > 0:
				data_date_list = list(data_dict.keys())
				if need_sort:
					data_date_list.sort()
				date_list = map(lambda x: x.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT), data_date_list)
				print ("%s" % (",".join(date_list)))
			print ("\n")


	def delete_database(self):
		# import pdb; pdb.set_trace()
		if self.xcfg["database_date"] is not None:
			# data_date = self.get_data_date(self.xcfg["database_date"])
			# print ("Delete Data on %s to the database..." % data_date.strftime(self.DEFAULT_DB_DATE_STRING_FORMAT))
			self.__delete_db(self.xcfg["database_date"])
		elif self.xcfg["database_all_date_range"]:
			find_res = self.__delete_db_range()
		elif self.xcfg["database_date_range"] is not None:
			self.__delete_db_range(self.xcfg["database_date_range_start"], self.xcfg["database_date_range_end"])


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


	@property
	def DatabaseDate(self):
		return self.xcfg["database_date"]


	@DatabaseDate.setter
	def DatabaseDate(self, database_date):
		self.xcfg["database_date"] = self.get_data_date(database_date)


	@property
	def SourceFolderpath(self):
		assert self.xcfg["source_folderpath"] is not None, "source_folderpath should NOT be NONE"
		return self.xcfg["source_folderpath"]


	@SourceFolderpath.setter
	def SourceFolderpath(self, source_folderpath):
		self.xcfg["source_folderpath"] = source_folderpath
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_filename"])
		# print ("SourceFolderpath: %s" % self.xcfg["source_filepath"])
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None


	@property
	def SourceFilename(self):
		assert self.xcfg["source_filename"] is not None, "source_filename should NOT be NONE"
		return self.xcfg["source_filename"]


	@SourceFilename.setter
	def SourceFilename(self, source_filename):
		self.xcfg["source_filename"] = source_filename
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_filename"])
		# print ("SourceFilename: %s" % self.xcfg["source_filepath"])
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None

# from dataclasses import dataclass, field # 記得要 import field
# import datetime
# @dataclass
# class Employee:
#     """Class that contains basic information about an employee."""
#     name: str
#     job: str
#     salary: int = 0
#     record_time: datetime.datetime = \
#         field(init=False, default_factory=datetime.datetime.now) # 資料紀錄時間
# # 創造實例的時候引數 *不可以* 包含 record_time，不然會出現 error


if __name__ == "__main__":

	cfg = {
		"database_date": "2022-10-07",
		"source_filename": "stock_chip_analysis@2022-10-07.xlsm",
	}
	with StockChipAnalysis(cfg) as obj:
		obj.check_data_source_difference(False)

	sys.exit(0)

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
	parser.add_argument('--list_analysis_method', required=False, action='store_true', help='List each analysis method and exit')
	parser.add_argument('--list_sheet_set_category', required=False, action='store_true', help='List each stock set and exit')
	parser.add_argument('--update_database', required=False, action='store_true', help='Update database and exit')
	parser.add_argument('--update_database_multiple', required=False, action='store_true', help='Update database from multiple XLS files and exit. Caution: The format of the XLS filename: {0}@20YY-mm-DD. Ex: {0}_2022-07-29'.format(StockChipAnalysis.DEFAULT_SOURCE_FILENAME))
	parser.add_argument('--find_database', required=False, action='store_true', help='Find database and exit')
	parser.add_argument('--delete_database', required=False, action='store_true', help='Delete database and exit')
	parser.add_argument('--list_database_date', required=False, action='store_true', help='List database date and exit')
	parser.add_argument('--database_date', required=False, help='The date of the data in the database. Ex: 2022-05-18. Caution: Update/Find/Delete Database')
	parser.add_argument('--database_date_range', required=False, help='The date range of the data in the database. Format: start_date,end_date. Ex: (1) 2022-05-18,2022-05-30 ; (2) 2022-05-18, ; (3) ,2022-05-30. Caution: Find/Delete Database')
	parser.add_argument('--database_all_date_range', required=False, action='store_true', help='The all date range of the data in the database')
	parser.add_argument('--create_report_by_sheet_set_category', required=False, help='Create a report by certain a sheet set category and exit')
	parser.add_argument('--check_data_source_difference', required=False, help='Check data source difference on a specific day and exit')
	parser.add_argument('-m', '--analysis_method', required=False, help='The method for chip analysis. Default: 0')	
	parser.add_argument('-d', '--show_detail', required=False, action='store_true', help='Show detailed data for each stock')
	parser.add_argument('-g', '--generate_report', required=False, action='store_true', help='Generate the report of the detailed data for each stock to the XLS file.')
	parser.add_argument('-r', '--report_filename', required=False, help='The filename of chip analysis report')
	parser.add_argument('-t', '--stock_list_filename', required=False, help='The filename of stock list for chip analysis')
	parser.add_argument('-l', '--stock_list', required=False, help='The list string of stock list for chip analysis. Ex: 2330,2317,2454,2308')
	parser.add_argument('--source_folderpath', required=False, help='Update database from the XLS files in the designated folder path. Ex: %s' % StockChipAnalysis.DEFAULT_SOURCE_FOLDERPATH)
	parser.add_argument('--source_filename', required=False, help='The filename of chip analysis data source')
	parser.add_argument('-c', '--sheet_set_category', required=False, help='The category for sheet set. Default: 0')	
	parser.add_argument('-n', '--need_all_sheet', required=False, action='store_true', help='The stock should be found in all sheets in the sheet name list')
	parser.add_argument('--search_history', required=False, action='store_true', help='The data source is from the database, otherwise from the excel file')
	parser.add_argument('-a', '--search_result_filename', required=False, help='The filename of stock list for search result')
	parser.add_argument('-o', '--output_search_result', required=False, action='store_true', help='Ouput the search result')
	parser.add_argument('-q', '--quiet', required=False, action='store_true', help="Don't print string on the screen")
	parser.add_argument('-s', '--sort', required=False, action='store_true', help="Show the data in order")
	parser.add_argument('-f', '--sort_limit', required=False, help="Limit the sorted data")
	# parser.add_argument('-f', '--update_datebase', required=False, help="Limit the sorted data")
	args = parser.parse_args()

	if args.list_analysis_method:
		help_str_list = [
			"Search data for the specific stocks from the file",
			"Search data for the specific stocks",
			"Search data for the whole stocks",
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
	if args.check_data_source_difference:
		StockChipAnalysis.check_data_source_difference(args.check_data_source_difference)
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
	if args.database_date is not None: cfg['database_date'] = args.database_date
	if args.database_date_range is not None: cfg['database_date_range'] = args.database_date_range
	if args.database_all_date_range: cfg['database_all_date_range'] = True
	if args.show_detail: cfg['show_detail'] = True
	if args.generate_report: cfg['generate_report'] = True
	if args.report_filename is not None: cfg['report_filename'] = args.report_filename
	if args.stock_list_filename is not None: cfg['stock_list_filename'] = args.stock_list_filename
	if args.stock_list is not None: cfg['stock_list'] = args.stock_list
	if args.source_folderpath is not None: cfg['source_folderpath'] = args.source_folderpath
	if args.source_filename is not None: cfg['source_filename'] = args.source_filename
	cfg['sheet_set_category'] = int(args.sheet_set_category) if args.sheet_set_category is not None else -1
	if args.need_all_sheet: cfg['need_all_sheet'] = True
	if args.report_filename is not None: cfg['report_filename'] = args.report_filename
	if args.search_history: cfg['search_history'] = True
	if args.search_result_filename is not None: cfg['search_result_filename'] = args.search_result_filename
	if args.output_search_result: cfg['output_search_result'] = True
	if args.quiet: cfg['quiet'] = True
	if args.sort: cfg['sort'] = True
	if args.sort_limit is not None: cfg['sort_limit'] = int(args.sort_limit)

	# import pdb; pdb.set_trace()
	with StockChipAnalysis(cfg) as obj:
		if args.update_database:
			 obj.update_database()
		elif args.update_database_multiple:
			fliename_dict = {}
			pattern = "%s.xlsm|%s@(20[\d]{2}-[\d]{2}-[\d]{2}).xlsm" % (StockChipAnalysis.DEFAULT_SOURCE_FILENAME, StockChipAnalysis.DEFAULT_SOURCE_FILENAME)
			# print (pattern)
			regex = re.compile(pattern)
			for filename in os.listdir(obj.SourceFolderpath):
				mobj = re.match(regex, filename)
				if mobj is None: continue
				# print (mobj.group(1))
				filepath = os.path.join(obj.SourceFolderpath, filename)
				if not os.path.isfile(filepath): continue
				file_date = obj.get_data_date(mobj.group(1))
				fliename_dict[file_date] = filename
				# print(f)
			# print (fliename_ordereddict)
			fliename_ordereddict = OrderedDict(sorted(fliename_dict.items(), key=lambda x: x[0]))
			for filedate, filename in fliename_ordereddict.items():
				print ("%s: %s" % (filedate, filename))
				# import pdb; pdb.set_trace()
				obj.DatabaseDate = filedate
				obj.SourceFilename = filename
				obj.update_database()
		elif args.find_database:
			obj.find_database()
		elif args.delete_database:
			obj.delete_database()
		elif args.list_database_date:
			obj.list_database_date()
		elif cfg['analysis_method'] == 0:
			obj.search_sheets_from_file()
		elif cfg['analysis_method'] == 1:
			obj.search_sheets()
		elif cfg['analysis_method'] == 2:
			obj.search_sheets(True)
		elif cfg['analysis_method'] == 3:
			obj.evaluate_bull_bear()
		else:
			raise ValueError("Incorrect Analysis Method Index")
