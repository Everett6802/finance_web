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
import xlrd
# import xlsxwriter
import argparse
# from datetime import datetime
# from pymongo import MongoClient
# from collections import OrderedDict
import csv
import collections


class ConvertibleBondAnalysis(object):

	DEFAULT_CB_FOLDERPATH =  "C:\\可轉債"
	DEFAULT_CB_SUMMARY_FILENAME = "可轉債總表"
	DEFAULT_CB_SUMMARY_FULL_FILENAME = "%s.csv" % DEFAULT_CB_SUMMARY_FILENAME
	DEFAULT_CB_QUOTATION_FILENAME = "可轉債報價"
	DEFAULT_CB_QUOTATION_FULL_FILENAME = "%s.xlsx" % DEFAULT_CB_QUOTATION_FILENAME
	DEFAULT_CB_QUOTATION_FIELD_TYPE = [str, str, float, float, int, float, float, str,]
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


	def __init__(self, cfg):
		self.xcfg = {
			"cb_folderpath": None,
			"cb_summary_filename": None,
			"cb_quotation_filename": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["cb_folderpath"] = self.DEFAULT_CB_FOLDERPATH if self.xcfg["cb_folderpath"] is None else self.xcfg["cb_folderpath"]
		self.xcfg["cb_summary_filename"] = self.DEFAULT_CB_SUMMARY_FULL_FILENAME if self.xcfg["cb_summary_filename"] is None else self.xcfg["cb_summary_filename"]
		self.xcfg["cb_summary_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_summary_filename"])
		self.xcfg["cb_quotation_filename"] = self.DEFAULT_CB_QUOTATION_FULL_FILENAME if self.xcfg["cb_quotation_filename"] is None else self.xcfg["cb_quotation_filename"]
		self.xcfg["cb_quotation_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_quotation_filename"])

		self.cb_summary = self.__read_cb_summary()
		self.workbook = None
		self.worksheet = None


	def __enter__(self):
# Open the workbook
		self.workbook = xlrd.open_workbook(self.xcfg["cb_quotation_filepath"])
		self.worksheet = self.workbook.sheet_by_index(0)
		return self


	def __exit__(self, type, msg, traceback):
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


	def __read_cb_summary(self):
		pattern = "(.+)\(([\d]{5,6})\)"
		cb_data = {}
		with open(self.xcfg["cb_summary_filepath"], newline='') as f:
			rows = csv.reader(f)
			regex = re.compile(pattern)
			title_list = None
			for index, row in enumerate(rows):
				# import pdb; pdb.set_trace()
				if index == 0: pass
				elif index == 1:
					title_list = list(map(lambda x: x.lstrip("=\"").rstrip("\"").rstrip("(%)"), row))
					# print(title_list)
					# CSVData = collections.namedtuple("CSVData", "%s" % (" ".join(title_list)))
				else:
					assert title_list is not None, "title_list should NOT be None"
					data_list = list(map(lambda x: x.lstrip("=\"").rstrip("\""), row))
					mobj = re.match(regex, data_list[0])
					if mobj is None: 
						raise ValueError("Incorrect format: %s" % data_list[0])
					data_list[0] = mobj.group(1)
					data_dict = dict(zip(title_list, data_list))
					cb_data[mobj.group(2)] = data_dict
				# print ("%s" % (",".join(data_list)))
		return cb_data


	def __read_cb_quotation(self):
		cb_data = {}
		# import pdb; pdb.set_trace()
		title_list = []
		for column_index in range(1, self.worksheet.ncols):
			title_value = self.worksheet.cell_value(0, column_index)
			title_list.append(title_value)
		for row_index in range(1, self.worksheet.nrows):
			data_key = self.worksheet.cell_value(row_index, 0)
			data_list = []
			data_key = self.worksheet.cell_value(row_index, 0)
			for column_index in range(1, self.worksheet.ncols):
				data_value = self.worksheet.cell_value(row_index, column_index)
				try:
					data_type = self.DEFAULT_CB_QUOTATION_FIELD_TYPE[column_index]
					data_value = data_type(data_value)
				except ValueError:
					# print "End row index: %d" % row_index
					data_value = None
					break
				# except Exception as e:
				# 	import pdb; pdb.set_trace()
				# 	print (e)
				data_list.append(data_value)
			data_dict = dict(zip(title_list, data_list))
			cb_data[data_key] = data_dict
		return cb_data


	def test(self):
		data_dict_summary = self.__read_cb_summary()
		# print (data_dict_summary)
		data_dict_quotation = self.__read_cb_quotation()
		# print (data_dict_quotation)
		if set(data_dict_summary.keys()) == set(data_dict_quotation.keys()):
			raise ValueError("The CB keys are NOT identical")


	# def __get_workbook(self):
	# 	if self.workbook is None:
	# 		# import pdb; pdb.set_trace()
	# 		self.workbook = xlrd.open_workbook(self.xcfg["cb_filepath"])
	# 		# print ("__get_workbook: %s" % self.xcfg["cb_filepath"])
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

