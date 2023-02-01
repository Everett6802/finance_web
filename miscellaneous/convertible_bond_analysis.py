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
from datetime import datetime
from datetime import date
import math
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


	@classmethod
	def __get_days(cls, end_date_str, start_date_str=None):
		# import pdb; pdb.set_trace()
		end_date = datetime.strptime(end_date_str,"%Y/%m/%d")
		start_date = datetime.strptime(start_date_str,"%Y/%m/%d") if start_date_str is not None else datetime.fromordinal(date.today().toordinal())
		days = int((end_date - start_date).days) + 1
		return days


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

		self.workbook = None
		self.worksheet = None
		self.cb_summary = self.__read_cb_summary()
		self.cb_id_list = list(self.cb_summary.keys())
		self.check_cb_id =  False


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
		# import pdb; pdb.set_trace()
		for row_index in range(1, self.worksheet.nrows):
			data_key = self.worksheet.cell_value(row_index, 0)
			data_list = []
			for column_index in range(1, self.worksheet.ncols):
				data_value = self.worksheet.cell_value(row_index, column_index)
				try:
					data_type = self.DEFAULT_CB_QUOTATION_FIELD_TYPE[column_index]
					data_value = data_type(data_value)
				except ValueError:
					# print "End row index: %d" % row_index
					data_value = None
				# except Exception as e:
				# 	import pdb; pdb.set_trace()
				# 	print (e)
				data_list.append(data_value)
			data_dict = dict(zip(title_list, data_list))
			cb_data[data_key] = data_dict
		if not self.check_cb_id:
			# import pdb; pdb.set_trace()
			cb_summary_id_set = set(self.cb_summary.keys())
			cb_quotation_id_set = set(cb_data.keys())
			cb_diff_id_set = cb_summary_id_set - cb_quotation_id_set
			if len(cb_diff_id_set) > 0:
				# raise ValueError("The CB keys are NOT identical: %s" % cb_diff_id_set)
				print("The CB IDs are NOT identical: %s" % cb_diff_id_set)
				for cb_id in list(cb_diff_id_set):
					self.cb_id_list.remove(cb_id)
			self.check_cb_id = True
		return cb_data


	def calculate_internal_rate_of_return(self, cb_quotation, use_percentage=True, filter_funcptr=None):
		irr_dict = {}
		# import pdb; pdb.set_trace()
		for cb_id in self.cb_id_list:
			cb_quotation_data = cb_quotation[cb_id]
			# print(cb_quotation_data)
			if cb_quotation_data["賣出一"] is None:
				continue
			days = self.__get_days(cb_quotation_data["到期日"])
			days_to_year = days / 365.0
			irr = math.pow(100.0 / cb_quotation_data["賣出一"], 1 / days_to_year) - 1
			if use_percentage:
				irr *= 100.0
			if filter_funcptr is not None:
				if not filter_funcptr(irr):
					continue
			irr_dict[cb_id] = {"商品": cb_quotation_data["商品"], "到期日": cb_quotation_data["到期日"], "年化報酬率": irr}
		return irr_dict


	def get_positive_internal_rate_of_return(self, cb_quotation, positive_threshold=0.1):
		filter_funcptr = lambda x: True if (x >= positive_threshold) else False
		return self.calculate_internal_rate_of_return(cb_quotation, use_percentage=True, filter_funcptr=filter_funcptr)  


	def test(self):
		# data_dict_summary = self.__read_cb_summary()
		# # print (data_dict_summary)
		data_dict_quotation = self.__read_cb_quotation()
		# print (data_dict_quotation)
		# print(self.calculate_internal_rate_of_return(data_dict_quotation))
		irr_dict = self.get_positive_internal_rate_of_return(data_dict_quotation)
		for irr_key, irr_data in irr_dict.items():
			print ("%s[%s]: %.2f  %s" % (irr_data["商品"], irr_key, float(irr_data["年化報酬率"]), irr_data["到期日"]))
		# print(irr_dict)


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

