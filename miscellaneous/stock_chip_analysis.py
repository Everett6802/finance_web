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
import argparse
from collections import OrderedDict


class StockChipAnalysis(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\Users\Price\Downloads" # os.getcwd()
	DEFAULT_SOURCE_FILENAME = "stock_chip_analysis.xlsm"
	DEFAULT_CONFIG_FOLDERPATH =  "C:\Users\Price\source"
	DEFAULT_CHIP_ANALYSIS_STOCK_LIST_FILENAME = "chip_analysis_stock_list.txt"
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
	SHEET_CATEGORY_DICT = {
		"consecutive_buy": [u"主力買超天數累計", u"法人買超天數累計", u"外資買超天數累計", u"投信買超天數累計",],
	}


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


	def __init__(self, cfg):
		self.xcfg = {
			"show_detail": False,
			"source_filepath": os.path.join(self.DEFAULT_SOURCE_FOLDERPATH, self.DEFAULT_SOURCE_FILENAME),
			"stock_list_filepath": os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.DEFAULT_CHIP_ANALYSIS_STOCK_LIST_FILENAME),
			"stock_list": None,
			# "buy_sheet_threshold": self.DEFAULT_BUY_SHEET_THRESHOLD,
			"sheet_name_list": None,
			"sheet_category": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		if cfg.has_key("sheet_category"):
			if cfg.has_key("sheet_name_list"):
				print "WARNING: The 'sheet_category' setting overwrite the 'sheet_name_list' one"
			if cfg.has_key("select_sheet_description_list"):
				print "WARNING: The 'sheet_category' setting overwrite the 'select_sheet_description_list' one"
			if cfg.has_key("buy_sheet_threshold"):
				print "WARNING: The 'sheet_category' setting overwrite the 'buy_sheet_threshold' one"
			self.xcfg["sheet_name_list"] = []
			for sheet_description in self.SHEET_CATEGORY_DICT[self.xcfg["sheet_category"]]:
				sheet_index = self.__get_sheet_index_by_description(sheet_description)
				if sheet_index == -1:
					raise RuntimeError("Unknown sheet description: %s" % sheet_description)
				self.xcfg["sheet_name_list"].append(sheet_index)		
			self.xcfg["buy_sheet_threshold"] = len(self.SHEET_CATEGORY_DICT[self.xcfg["sheet_category"]])
		else:
			if cfg.has_key("select_sheet_description_list"):
				if cfg.has_key("sheet_name_list"):
					print "WARNING: The 'select_sheet_description_list' setting overwrite the 'sheet_name_list' one"
				self.xcfg["sheet_name_list"] = []
				for sheet_description in self.xcfg["select_sheet_description_list"]:
					sheet_index = self.__get_sheet_index_by_description(sheet_description)
					if sheet_index == -1:
						raise RuntimeError("Unknown sheet description: %s" % sheet_description)
					self.xcfg["sheet_name_list"].append(sheet_index)

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
		return data_dict


	def __read_sheet_title_bar(self, sheet_name):
		# import pdb; pdb.set_trace()
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
		return title_bar_list


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


	def __find_sheet_occurrence(self, ignore_sheet_func_ptr=None, sheet_data_func_ptr=None):
		stock_number_sheet_dict = {}
		stock_number_extra_dict = {}
		# import pdb; pdb.set_trace()
		if self.xcfg["sheet_name_list"] is None:
			self.xcfg["sheet_name_list"] = self.DEFAULT_SHEET_NAME_LIST
		for sheet_index in self.xcfg["sheet_name_list"]:
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


	def __search_stock_sheets(self):
		# import pdb; pdb.set_trace()
		sheet_data_func_ptr = (lambda x: x) if self.xcfg["show_detail"] else (lambda x: x[0])
		sheet_data_collection_dict = self.__collect_sheet_data(sheet_data_func_ptr)
		no_data = True
		for stock_number in self.xcfg["stock_list"]:
			if not sheet_data_collection_dict.has_key(stock_number):
				continue
			no_data = False
			stock_sheet_data_collection_dict = sheet_data_collection_dict[stock_number]
			if self.xcfg["show_detail"]:
				stock_name = stock_sheet_data_collection_dict.values()[0][0]
				print "=== %s(%s) ===" % (stock_number, stock_name)
				for sheet_name, sheet_data_list in stock_sheet_data_collection_dict.items():
					sheet_title_bar_list = self.__read_sheet_title_bar(sheet_name)
					sheet_data_list_len = len(sheet_data_list)
					sheet_title_bar_list_len = len(sheet_title_bar_list)
					assert sheet_data_list_len == sheet_title_bar_list_len, "The list lengths are NOT identical, sheet_data_list_len: %d, sheet_title_bar_list_len: %d" % (sheet_data_list_len, sheet_title_bar_list_len)
					print "* %s" % sheet_name
					print "%s" % ",".join(["%s[%s]" % elem for elem in zip(sheet_title_bar_list[1:], sheet_data_list[1:])])
			else:
				stock_name = stock_sheet_data_collection_dict.values()[0]
				print "=== %s(%s) ===" % (stock_number, stock_name)
				print "%s" % (u",".join([stock_sheet_data_key for stock_sheet_data_key in stock_sheet_data_collection_dict.keys()]))
		if no_data: print "*** No Data ***"	


	def search_sheets_from_file(self):
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(self.xcfg['stock_list_filepath']):
			raise RuntimeError("The file[%s] does NOT exist" % self.xcfg['stock_list_filepath'])
		self.xcfg["stock_list"] = []
		with open(self.xcfg['stock_list_filepath'], 'r') as fp:
			for line in fp:
				self.xcfg["stock_list"].append(line.strip("\n"))
		self.__search_stock_sheets()


	def search_sheets(self):
		if self.xcfg['stock_list'] is None:
			raise RuntimeError("The search target list should NOT be None")
		self.xcfg['stock_list'] = self.xcfg['stock_list'].split(",")
		self.__search_stock_sheets()


	def search_buy(self):
		# import pdb; pdb.set_trace()
		sheet_occurrence_dict, sheet_occurrence_extra_dict = self.__find_sheet_occurrence(lambda x: self.SHEET_METADATA_LIST[x]["direction"] == '-', lambda x: x[0])
		filtered_sheet_occurrence_dict = dict(filter(lambda x: len(x[1]) >= self.xcfg["buy_sheet_threshold"], sheet_occurrence_dict.items()))
		filtered_sheet_occurrence_ordereddict = OrderedDict(sorted(filtered_sheet_occurrence_dict.items(), key=lambda x: x[1]))
		for stock_number, sheet_name_list in filtered_sheet_occurrence_ordereddict.items():
			print "=== %s(%s) ===" % (stock_number, sheet_occurrence_extra_dict[stock_number])
			print "%s" % (u",".join([self.SHEET_METADATA_LIST[index]["description"] for index in sheet_name_list]))


	@property
	def StockList(self):
		return self.xcfg["stock_list"]


	@StockList.setter
	def StockList(self, stock_list):
		self.xcfg["stock_list"] = stock_list


if __name__ == "__main__":
	
	help_str_list = [
		"Search sheet for each stock from the file",
		"Search sheet for each stock",
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
	parser.add_argument('-s', '--show_detail', required=False, action='store_true', help='Show detailed data for each stock')
	parser.add_argument('-f', '--stock_list_filepath', required=False, help='The filepath of stock list for chip analysis')
	parser.add_argument('-l', '--stock_list', required=False, help='The list string of stock list for chip analysis. Ex: 2330,2317,2454,2308')
	parser.add_argument('-i', '--sheet_name_list', required=False, help='The sheet index for searching')
	parser.add_argument('-b', '--buy_sheet_threshold', required=False, help='The threshold of the sheet count that institutional investors/large trader buy')
	parser.add_argument('-d', '--select_sheet_description_list', required=False, help='Select the sheet description for searching')
	parser.add_argument('--select_sheet_category_consecutive_buy', required=False, action='store_true', help='Select the sheet category: consecutive_buy')
	args = parser.parse_args()
	# import pdb; pdb.set_trace()

	cfg = {}
	cfg['analysis_method'] = int(args.analysis_method) if args.analysis_method is not None else 0
	if args.show_detail: cfg['show_detail'] = True
	if args.stock_list_filepath is not None: cfg['stock_list_filepath'] = args.stock_list_filepath
	if args.stock_list is not None: cfg['stock_list'] = args.stock_list
	if args.sheet_name_list is not None: cfg['sheet_name_list'] = args.sheet_name_list
	if args.buy_sheet_threshold is not None: cfg['buy_sheet_threshold'] = int(args.buy_sheet_threshold)
	if args.select_sheet_description_list is not None: cfg['select_sheet_description_list'] = args.select_sheet_description_list
	if args.select_sheet_category_consecutive_buy: cfg['sheet_category'] = "consecutive_buy"
		
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
