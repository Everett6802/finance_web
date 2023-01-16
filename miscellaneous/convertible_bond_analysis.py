#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import errno
# '''
# Question: How to Solve xlrd.biffh.XLRDError: Excel xlsx file; not supported ?
# Answer : The latest version of xlrd(2.01) only supports .xls files. Installing the older version 1.2.0 to open .xlsx files.
# '''
# import xlrd
# import xlsxwriter
import argparse
# from datetime import datetime
# from pymongo import MongoClient
# from collections import OrderedDict
import csv
import collections


class ConvertibleBondAnalysis(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\\可轉債"
	DEFAULT_SOURCE_SUMMARY_FILENAME = "可轉債總表"
	DEFAULT_SOURCE_SUMMARY_FULL_FILENAME = "%s.csv" % DEFAULT_SOURCE_SUMMARY_FILENAME
	DEFAULT_SOURCE_QUOTATION_FILENAME = "可轉債報價"
	DEFAULT_SOURCE_QUOTATION_FULL_FILENAME = "%s.xlsx" % DEFAULT_SOURCE_QUOTATION_FILENAME
	# DEFAULT_CONFIG_FOLDERPATH =  "C:\\Users\\%s\\source" % os.getlogin()
	# DEFAULT_STOCK_LIST_FILENAME = "chip_analysis_stock_list.txt"
	# DEFAULT_REPORT_FILENAME = "chip_analysis_report.xlsx"
	# DEFAULT_SEARCH_RESULT_FILENAME = "search_result_stock_list.txt"


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
	def __read_from_csv(cls, filepath):
		pattern = "(.+)\(([\d]{5,6})\)"
		csv_data = {}
		with open(filepath, newline='') as f:
			rows = csv.reader(f)
			regex = re.compile(pattern)
			for index, row in enumerate(rows):
				# import pdb; pdb.set_trace()
				if index == 0: pass
				elif index == 1:
					title_list = list(map(lambda x: x.lstrip("=\"").rstrip("\"").rstrip("(%)"), row))
					print(title_list)
					CSVData = collections.namedtuple("CSVData", "%s" % (" ".join(title_list)))
				else:
					data_list = list(map(lambda x: x.lstrip("=\"").rstrip("\""), row))
					mobj = re.match(regex, data_list[0])
					if mobj is None: 
						raise ValueError("Incorrect format: %s" % data_list[0])
					data_list[0] = mobj.group(1)
					csv_data[mobj.group(2)] = data_list
				# print ("%s" % (",".join(data_list)))
		return csv_data


	def __read_data(self):
		data_dict = {}
		dt_now = datetime.datetime.now()
		print "Read %s at %s" % (os.path.basename(self.xcfg["source_filepath"]), dt_now.strftime(self.DATETIME_FORMAT_STR))
		# import pdb; pdb.set_trace()
		for row_index in range(self.DATA_CELL_ROW_START_INDEX, self.DATA_CELL_ROW_END_INDEX):
			key = None
			try:
				key = self.worksheet.cell_value(row_index, 0)
			except IndexError:
				# print "End row index: %d" % row_index
				break
			# print "row_index: %d, %s" % (row_index, self.worksheet.cell_value(row_index, 0))
			data_dict[key] = {}
			data_cell_column_list = None
			column_index_data = self.DATA_CELL_ROW_COLUMN_INDEX_DICT.get(row_index, None)
			if column_index_data is None:
				data_cell_column_list = range(self.DATA_CELL_COLUMN_DEF_START_INDEX, self.DATA_CELL_COLUMN_DEF_END_INDEX)
			else:
				if type(column_index_data) is dict:
					data_cell_column_list = range(self.DATA_CELL_ROW_COLUMN_INDEX_DICT[row_index]["from"], self.DATA_CELL_ROW_COLUMN_INDEX_DICT[row_index]["to"])
				elif type(column_index_data) is list:
					data_cell_column_list = self.DATA_CELL_ROW_COLUMN_INDEX_DICT[row_index]
				else:
					raise RuntimeError("Unknown column index range in row: %d" % row_index)
			# import pdb; pdb.set_trace()
			for column_index in data_cell_column_list:
				# print "row: %d, column: %d" % (row_index, column_index)
				cell_value = self.worksheet.cell_value(row_index, column_index)
				# print "value: %s" % str(cell_value)
# Check if this option is traded
				if self.__is_string(cell_value):
					data_dict[key][self.DATA_CELL_COLUMN_TITLE_NAME_LIST[column_index]] = None
				# print "%s %d %d" % (key, row_index, column_index)
				# import pdb; pdb.set_trace()
				elif self.DATA_CELL_COLUMN_TYPE_LIST[column_index] == "float":
					data_dict[key][self.DATA_CELL_COLUMN_TITLE_NAME_LIST[column_index]] = float(cell_value)
				elif self.DATA_CELL_COLUMN_TYPE_LIST[column_index] == "int":
					data_dict[key][self.DATA_CELL_COLUMN_TITLE_NAME_LIST[column_index]] = int(cell_value)
				else:
					raise ValueError("Unknown type: %s" % self.DATA_CELL_COLUMN_TYPE_LIST[index])
		# import pdb; pdb.set_trace()
		data_dict["created_at"] = dt_now
		return data_dict


	def __init__(self, cfg):
		self.xcfg = {
			"source_folderpath": None,
			"source_summary_filename": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["source_folderpath"] = self.DEFAULT_SOURCE_FOLDERPATH if self.xcfg["source_folderpath"] is None else self.xcfg["source_folderpath"]
		self.xcfg["source_summary_filename"] = self.DEFAULT_SOURCE_SUMMARY_FULL_FILENAME if self.xcfg["source_summary_filename"] is None else self.xcfg["source_summary_filename"]
		self.xcfg["source_summary_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_summary_filename"])
		self.xcfg["source_quotation_filename"] = self.DEFAULT_SOURCE_QUATATION_FULL_FILENAME if self.xcfg["source_quotation_filename"] is None else self.xcfg["source_quotation_filename"]
		self.xcfg["source_quotation_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_quotation_filename"])


	def __enter__(self):
		# Open the workbook
		# self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		# if self.xcfg["output_search_result"]:
		# 	self.search_result_txtfile = open(self.xcfg["search_result_filepath"], "w")
		return self


	def __exit__(self, type, msg, traceback):
		# if self.workbook is not None:
		# 	self.workbook.release_resources()
		# 	del self.workbook
		# 	self.workbook = None
		return False


	def test(self):
		self.__read_from_csv(self.xcfg["source_summary_filepath"])


	# def __get_workbook(self):
	# 	if self.workbook is None:
	# 		# import pdb; pdb.set_trace()
	# 		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
	# 		# print ("__get_workbook: %s" % self.xcfg["source_filepath"])
	# 	return self.workbook


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
	parser.add_argument('--list_analysis_method', required=False, action='store_true', help='List each analysis method and exit')
	args = parser.parse_args()

	cfg = {
	}
	with ConvertibleBondAnalysis(cfg) as obj:
		obj.test()
