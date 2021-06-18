#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import errno
# import requests
'''
Question: How to Solve xlrd.biffh.XLRDError: Excel xlsx file; not supported ?
Answer : The latest version of xlrd(2.01) only supports .xls files. Installing the older version 1.2.0 to open .xlsx files.
'''
import xlrd
import argparse
from collections import OrderedDict
# import time
# import json
# import datetime


class StockChipAnalysis(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\Users\Price\Downloads" # os.getcwd()
	DEFAULT_SOURCE_FILENAME = "stock_chip_analysis.xlsx"
	DEFAULT_CONFIG_FOLDERPATH =  "C:\Users\Price\source"
	DEFAULT_SEARCH_STOCK_SHEET_FILENAME = "search_sheet_stock_list.txt"
	SHEET_METADATA_LIST = [
		{ # Dummy
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "法人共同買超累計",
			# "column_length": 10,
			"direction": "+",
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "主力買超天數累計",
			# "column_length": 15,
			"direction": "+",
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "法人買超天數累計",
			# "column_length": 15,
			"direction": "+",
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "外資買超天數累計",
			# "column_length": 15,
			"direction": "+",
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "投信買超天數累計",
			# "column_length": 15,
			"direction": "+",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "外資買最多股",
			# "column_length": 7,
			"direction": "+",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "外資賣最多股",
			# "column_length": 7,
			"direction": "-",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "投信買最多股",
			# "column_length": 7,
			"direction": "+",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "投信賣最多股",
			# "column_length": 7,
			"direction": "-",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "主力買最多股",
			# "column_length": 7,
			"direction": "+",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "主力賣最多股",
			# "column_length": 7,
			"direction": "-",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "籌碼排行-買超金額",
			# "column_length": 13,
			"direction": "+",
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "籌碼排行-賣超金額",
			# "column_length": 13,
			"direction": "-",
		},
	]
	SHEET_METADATA_LIST_LEN = len(SHEET_METADATA_LIST)
	# import pdb; pdb.set_trace()
	DEFAULT_BUY_SHEET_THRESHOLD = len([sheet_metadata for sheet_metadata in SHEET_METADATA_LIST[1:] if sheet_metadata["direction"] == "+"]) - 2


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


	def __init__(self, cfg):
		self.xcfg = {
			"show_detail": False,
			"source_filepath": os.path.join(self.DEFAULT_SOURCE_FOLDERPATH, self.DEFAULT_SOURCE_FILENAME),
			"search_sheet_filepath": os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.DEFAULT_SEARCH_STOCK_SHEET_FILENAME),
			"search_sheet_list": None,
			"buy_sheet_threshold": self.DEFAULT_BUY_SHEET_THRESHOLD,
		}
		self.xcfg.update(cfg)

		self.workbook = None
		self.sheet_title_bar_list = None


	def __enter__(self):
		# Open the workbook
		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		return self


	def __exit__(self, type, msg, traceback):
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


	def __read_sheet_data(self, sheet_index):
		# import pdb; pdb.set_trace()
		sheet_metadata = self.SHEET_METADATA_LIST[sheet_index]
		# print u"Read sheet: %s" % sheet_metadata["description"].decode("utf8")
		assert self.workbook is not None, "self.workbook should NOT be None"
		worksheet = self.workbook.sheet_by_index(sheet_index)
		# https://www.itread01.com/content/1549650266.html
		# print sheet1.name,sheet1.nrows,sheet1.ncols    #Sheet1 6 4
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
		return data_dict


	def __read_sheet_title_bar(self, sheet_index):
		# import pdb; pdb.set_trace()
		sheet_metadata = self.SHEET_METADATA_LIST[sheet_index]
		worksheet = self.workbook.sheet_by_index(sheet_index)
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
		return title_bar_list


	def __find_sheet_index(self, stock_number, sheet_data_func_ptr=None):
		sheet_index_list = []
		sheet_index_data_list = []
		for sheet_index in range(1, self.SHEET_METADATA_LIST_LEN):
			data_dict = self.__read_sheet_data(sheet_index)
			# print data_dict
			if data_dict.has_key(stock_number):
				sheet_index_list.append(sheet_index)
				if sheet_data_func_ptr is not None:
					sheet_index_data = sheet_data_func_ptr(data_dict[stock_number])
					sheet_index_data_list.append(sheet_index_data)
		return sheet_index_list, sheet_index_data_list


	def __find_sheet_occurrence(self, ignore_sheet_func_ptr=None, sheet_data_func_ptr=None):
		stock_number_sheet_dict = {}
		stock_number_extra_dict = {}
		for sheet_index in range(1, self.SHEET_METADATA_LIST_LEN):
			if ignore_sheet_func_ptr is not None and ignore_sheet_func_ptr(sheet_index):
				continue
			data_dict = self.__read_sheet_data(sheet_index)
			for stock_number, stock_data in data_dict.items():
				if stock_number_sheet_dict.has_key(stock_number):
					# stock_number_sheet_dict[stock_number] = stock_number_sheet_dict[stock_number] + 1
					stock_number_sheet_dict[stock_number].append(sheet_index)					
				else:
					# stock_number_sheet_dict[stock_number] = 1
					stock_number_sheet_dict[stock_number] = [sheet_index,]
					if sheet_data_func_ptr is not None:
						stock_number_extra_dict[stock_number] = sheet_data_func_ptr(stock_data)
		return stock_number_sheet_dict, stock_number_extra_dict


	def search_sheets_from_file(self):
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(self.xcfg['search_sheet_filepath']):
			raise RuntimeError("The file[%s] does NOT exist" % self.xcfg['search_sheet_filepath'])
		no_data = True
		with open(self.xcfg['search_sheet_filepath'], 'r') as fp:
			sheet_data_func_ptr = (lambda x: x) if self.xcfg["show_detail"] else (lambda x: x[0])
			for line in fp:
				stock_number = line.strip("\n")
				sheet_index_list, sheet_index_data_list = obj.__find_sheet_index(stock_number, sheet_data_func_ptr)
				sheet_index_list_len = len(sheet_index_list)
				sheet_index_data_list_len = len(sheet_index_data_list)
				assert sheet_index_list_len == sheet_index_data_list_len, "The list lengths are NOT identical, sheet_index_list_len: %d, sheet_index_data_list_len: %d" % (sheet_index_list_len, sheet_index_data_list_len) 
				if sheet_index_list_len != 0:
					no_data = False
					if self.xcfg["show_detail"]:
						print "=== %s(%s) ===" % (stock_number, sheet_index_data_list[0][0])
						for i in range(sheet_index_list_len):
							sheet_index = sheet_index_list[i]
							sheet_index_data = sheet_index_data_list[i]
							sheet_title_bar = self.__read_sheet_title_bar(sheet_index)
							sheet_index_data_len = len(sheet_index_data)
							sheet_title_bar_len = len(sheet_title_bar)
							assert sheet_index_data_len == sheet_title_bar_len, "The list lengths are NOT identical, sheet_index_data_len: %d, sheet_title_bar_len: %d" % (sheet_index_list_len, sheet_title_bar_len)
							# import pdb; pdb.set_trace()
							print "* %s" % self.SHEET_METADATA_LIST[sheet_index]["description"].decode("utf8")
							print "%s" % ",".join(["%s[%s]" % (sheet_title_bar[j], sheet_index_data[j]) for j in range(1, sheet_index_data_len)])
					else:
						print "=== %s(%s) ===" % (stock_number, sheet_index_data_list[0])
						print "%s" % (u",".join([self.SHEET_METADATA_LIST[index]["description"].decode("utf8") for index in sheet_index_list]))
		if no_data: print "*** No Data ***"


	def search_sheets(self):
		if self.xcfg['search_sheet_list'] is None:
			raise RuntimeError("The search target list should NOT be None")
		stock_number_list = self.xcfg['search_sheet_list'].split(",")
		no_data = True
		sheet_data_func_ptr = (lambda x: x) if self.xcfg["show_detail"] else (lambda x: x[0])
		# import pdb; pdb.set_trace()
		for stock_number in stock_number_list:
			sheet_index_list, sheet_index_data_list = obj.__find_sheet_index(stock_number, sheet_data_func_ptr)
			sheet_index_list_len = len(sheet_index_list)
			sheet_index_data_list_len = len(sheet_index_data_list)
			if sheet_index_list_len != 0:
				no_data = False
				if self.xcfg["show_detail"]:
					print "=== %s(%s) ===" % (stock_number, sheet_index_data_list[0][0])
					for i in range(sheet_index_list_len):
						sheet_index = sheet_index_list[i]
						sheet_index_data = sheet_index_data_list[i]
						sheet_title_bar = self.__read_sheet_title_bar(sheet_index)
						sheet_index_data_len = len(sheet_index_data)
						sheet_title_bar_len = len(sheet_title_bar)
						assert sheet_index_data_len == sheet_title_bar_len, "The list lengths are NOT identical, sheet_index_data_len: %d, sheet_title_bar_len: %d" % (sheet_index_list_len, sheet_title_bar_len)
						# import pdb; pdb.set_trace()
						print "* %s" % self.SHEET_METADATA_LIST[sheet_index]["description"].decode("utf8")
						print "%s" % ",".join(["%s[%s]" % (sheet_title_bar[j], sheet_index_data[j]) for j in range(1, sheet_index_data_len)])
				else:
					print "=== %s(%s) ===" % (stock_number, sheet_index_data_list[0])
					print "%s" % (u",".join([self.SHEET_METADATA_LIST[index]["description"].decode("utf8") for index in sheet_index_list]))
		if no_data: print "*** No Data ***"


	def search_buy(self):
		sheet_occurrence_dict, sheet_occurrence_extra_dict = self.__find_sheet_occurrence(lambda x: self.SHEET_METADATA_LIST[x]["direction"] == '-', lambda x: x[0])
		filtered_sheet_occurrence_dict = dict(filter(lambda x: len(x[1]) >= self.xcfg["buy_sheet_threshold"], sheet_occurrence_dict.items()))
		filtered_sheet_occurrence_ordereddict = OrderedDict(sorted(filtered_sheet_occurrence_dict.items(), key=lambda x: x[1]))
		for stock_number, sheet_index_list in filtered_sheet_occurrence_ordereddict.items():
			print "=== %s(%s) ===\n%s" % (stock_number, sheet_occurrence_extra_dict[stock_number], u",".join([self.SHEET_METADATA_LIST[index]["description"].decode("utf8") for index in sheet_index_list]))


if __name__ == "__main__":
	
	help_str_list = [
		"Search sheet index for each stock from the file",
		"Search sheet index for each stock",
		"Search stocks which institutional investors/large trader buy",
	]
	help_str_list_len = len(help_str_list)
	print "************ Analysis Method ************"
	for index, help_str in enumerate(help_str_list):
		print "%d  %s" % (index, help_str)
	print "*****************************************"

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
	parser.add_argument('-m', '--analysis_method', required=False, help='The method for chip analysis. Default: 0')	
	parser.add_argument('-d', '--show_detail', required=False, action='store_true', help='Show detailed data for each stock')
	parser.add_argument('-f', '--search_sheet_filepath', required=False, help='The filepath of stock list for searching for sheet')
	parser.add_argument('-l', '--search_sheet_list', required=False, help='The list string of stock list for searching for sheet. Ex: 2330,2317,2454,2308')
	parser.add_argument('-b', '--buy_sheet_threshold', required=False, help='The threshold of the sheet count that institutional investors/large trader buy')

# 	# parser.add_argument('-d', '--disable_check_time', required=False, action='store_true', help='No need to check time for collecting data')
# 	parser.add_argument('-o', '--one_shot_query', required=False, action='store_true', help='Collect data immediately')
# 	parser.add_argument('-s', '--start_time', required=False, help='The start time of collecting data. Format: HH:mm')
# 	parser.add_argument('-e', '--end_time', required=False, help='The end_time of collecting data. Format: HH:mm')
	args = parser.parse_args()
	# import pdb; pdb.set_trace()
	cfg = {}
	cfg['analysis_method'] = int(args.analysis_method) if args.analysis_method is not None else 0
	if args.show_detail: cfg['show_detail'] = True
	if args.search_sheet_filepath is not None: cfg['search_sheet_filepath'] = args.search_sheet_filepath
	if args.search_sheet_list is not None: cfg['search_sheet_list'] = args.search_sheet_list
	if args.buy_sheet_threshold is not None: cfg['buy_sheet_threshold'] = int(args.buy_sheet_threshold)
		
	# import pdb; pdb.set_trace()
	with StockChipAnalysis(cfg) as obj:
		if cfg['analysis_method'] == 0:
			obj.search_sheets_from_file()
		elif cfg['analysis_method'] == 1:
			obj.search_sheets()
		elif cfg['analysis_method'] == 2:
			obj.search_buy()
		else:
			raise ValueError("Analysis Method Index should be in the range [0, %d)" % help_str_list_len)


