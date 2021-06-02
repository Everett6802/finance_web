#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
# import requests
'''
Question: How to Solve xlrd.biffh.XLRDError: Excel xlsx file; not supported ?
Answer : The latest version of xlrd(2.01) only supports .xls files. Installing the older version 1.2.0 to open .xlsx files.
'''
import xlrd
import argparse
# import time
# import json
# import datetime


class StockChipAnalysis(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\Users\Price\Downloads" # os.getcwd()
	DEFAULT_SOURCE_FILENAME = "stock_chip_analysis.xlsx"
	DEFAULT_CONFIG_FOLDERPATH =  "C:\Users\Price\source"
	DEFAULT_SEARCH_TARGET_FILENAME = "search_target_list.txt"
	SHEET_METADATA_LIST = [
		{ # Dummy
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "法人共同買超累計",
			"column_length": 10,
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "法人買超天數累計",
			"column_length": 15,
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "外資買超天數累計",
			"column_length": 15,
		},
		{
			"key_mode": 0, # 1476.TW
			"description": "投信買超天數累計",
			"column_length": 15,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "外資買最多股",
			"column_length": 7,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "外資最賣多股",
			"column_length": 7,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "投信買最多股",
			"column_length": 7,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "投信最賣多股",
			"column_length": 7,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "主力買最多股",
			"column_length": 7,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "主力最賣多股",
			"column_length": 7,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "籌碼排行-買超金額",
			"column_length": 13,
		},
		{
			"key_mode": 1, # 陽明(2609)
			"description": "籌碼排行-賣超金額",
			"column_length": 13,
		},
	]
	SHEET_METADATA_LIST_LEN = len(SHEET_METADATA_LIST)

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
			"source_filepath": os.path.join(self.DEFAULT_SOURCE_FOLDERPATH, self.DEFAULT_SOURCE_FILENAME),
			"search_target_filepath": os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.DEFAULT_SEARCH_TARGET_FILENAME),
			"search_target_list": None,
		}
		self.xcfg.update(cfg)

		self.workbook = None


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


	def __read_data(self, sheet_index):
		# import pdb; pdb.set_trace()
		sheet_metadata = self.SHEET_METADATA_LIST[sheet_index]
		# print u"Read sheet: %s" % sheet_metadata["description"].decode("utf8")
		assert self.workbook is not None, "self.workbook should NOT be None"
		worksheet = self.workbook.sheet_by_index(sheet_index)
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
			for column_index in range(1, sheet_metadata["column_length"]):
				data_dict[stock_number].append(worksheet.cell_value(row_index, column_index))
			row_index += 1
			# print "%d -- %s" % (row_index, stock_number)
		return data_dict


	def __find_sheet_index(self, stock_number, sheet_data_func_ptr=None):
		sheet_index_list = []
		sheet_index_data_list = []
		for sheet_index in range(1, self.SHEET_METADATA_LIST_LEN):
			data_dict = self.__read_data(sheet_index)
			# print data_dict
			if data_dict.has_key(stock_number):
				sheet_index_list.append(sheet_index)
				if sheet_data_func_ptr is not None:
					sheet_index_data = sheet_data_func_ptr(data_dict[stock_number])
					sheet_index_data_list.append(sheet_index_data)
		return sheet_index_list, sheet_index_data_list


	def search_targets_from_file(self):
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(self.xcfg['search_target_filepath']):
			raise RuntimeError("The file[%s] does NOT exist" % self.xcfg['search_target_filepath'])
		no_data = True
		with open(self.xcfg['search_target_filepath'], 'r') as fp:
			for line in fp:
				stock_number = line.strip("\n")
				sheet_index_list, sheet_index_data_list = obj.__find_sheet_index(stock_number, lambda x: x[0])
				if len(sheet_index_list) != 0:
					no_data = False
					print "=== %s(%s) ===\n%s" % (stock_number, sheet_index_data_list[0], u",".join([self.SHEET_METADATA_LIST[index]["description"].decode("utf8") for index in sheet_index_list]))
		if no_data: print "*** No Data ***"


	def search_targets(self):
		if self.xcfg['search_target_list'] is None:
			raise RuntimeError("The search target list should NOT be None")
		stock_number_list = self.xcfg['search_target_list'].split(",")
		no_data = True
		for stock_number in stock_number_list:
			sheet_index_list, sheet_index_data_list = obj.__find_sheet_index(stock_number, lambda x: x[0])
			if len(sheet_index_list) != 0:
				no_data = False
				print "=== %s(%s) ===\n%s" % (stock_number, sheet_index_data_list[0], u",".join([self.SHEET_METADATA_LIST[index]["description"].decode("utf8") for index in sheet_index_list]))
		if no_data: print "*** No Data ***"


if __name__ == "__main__":
	
	help_str_list = [
		"Search target from the file",
		"Input search target",
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
	parser.add_argument('-m', '--analysis_method', required=False, help='The method for chip analysis. Default: 0')	
	parser.add_argument('-f', '--search_target_filepath', required=False, help='The filepath of search target list. Default: ')
	parser.add_argument('-l', '--search_target_list', required=False, help='Search target list string. Ex: 2330,2317,2454,2308')
# How to add option without any argument? use action='store_true'
	'''
	'store_true' and 'store_false' - 这些是 'store_const' 分别用作存储 True 和 False 值的特殊用例。另外，它们的默认值分别为 False 和 True。例如:

	>>> parser = argparse.ArgumentParser()
	>>> parser.add_argument('--foo', action='store_true')
	>>> parser.add_argument('--bar', action='store_false')
	>>> parser.add_argument('--baz', action='store_false')
    '''
# 	# parser.add_argument('-d', '--disable_check_time', required=False, action='store_true', help='No need to check time for collecting data')
# 	parser.add_argument('-o', '--one_shot_query', required=False, action='store_true', help='Collect data immediately')
# 	parser.add_argument('-s', '--start_time', required=False, help='The start time of collecting data. Format: HH:mm')
# 	parser.add_argument('-e', '--end_time', required=False, help='The end_time of collecting data. Format: HH:mm')
	args = parser.parse_args()
	cfg = {}
	cfg['analysis_method'] = int(args.analysis_method) if args.analysis_method is not None else 0
	if args.search_target_filepath is not None: cfg['search_target_filepath'] = args.search_target_filepath
	if args.search_target_list is not None: cfg['search_target_list'] = args.search_target_list
		
	# import pdb; pdb.set_trace()
	with StockChipAnalysis(cfg) as obj:
		if cfg['analysis_method'] == 0:
			obj.search_targets_from_file()
		elif cfg['analysis_method'] == 1:
			obj.search_targets()
		else:
			raise ValueError("Analysis Method Index should be in the range [0, %d)" % help_str_list_len)


