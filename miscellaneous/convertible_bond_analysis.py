#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import errno
import traceback
# '''
# Question: How to Solve xlrd.biffh.XLRDError: Excel xlsx file; not supported ?
# Answer : The latest version of xlrd(2.01) only supports .xls files. Installing the older version 1.2.0 to open .xlsx files.
# '''
import xlrd
# import xlsxwriter
import argparse
from datetime import datetime
from datetime import date
from datetime import timedelta
import math
import csv
import collections


class ConvertibleBondAnalysis(object):

	DEFAULT_CB_FOLDERPATH =  "C:\\可轉債"
	DEFAULT_CB_SUMMARY_FILENAME = "可轉債總表"
	DEFAULT_CB_SUMMARY_FULL_FILENAME = "%s.csv" % DEFAULT_CB_SUMMARY_FILENAME
# ['可轉債商品', '到期日', '可轉換日', '票面利率', '上次付息日', '轉換價格', '現股收盤價', '可轉債價格', '套利報酬', '年化殖利率', '']
	DEFAULT_CB_SUMMARY_FIELD_TYPE = [str, str, str, float, str, float, float, float, float, float, str,]
	DEFAULT_CB_PUBLISH_FILENAME = "可轉債發行"
	DEFAULT_CB_PUBLISH_FULL_FILENAME = "%s.csv" % DEFAULT_CB_PUBLISH_FILENAME
# ['債券簡稱', '發行人', '發行日期', '到期日期', '年期', '發行總面額', '發行資料']
	DEFAULT_CB_PUBLISH_FIELD_TYPE = [str, str, str, str, int, int, str,]
	DEFAULT_CB_QUOTATION_FILENAME = "可轉債報價"
	DEFAULT_CB_QUOTATION_FULL_FILENAME = "%s.xlsx" % DEFAULT_CB_QUOTATION_FILENAME
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '到期日']
	DEFAULT_CB_QUOTATION_FIELD_TYPE = [str, float, float, int, float, float, str,]
	DEFAULT_CB_STOCK_QUOTATION_FILENAME = "可轉債個股報價"
	DEFAULT_CB_STOCK_QUOTATION_FULL_FILENAME = "%s.xlsx" % DEFAULT_CB_STOCK_QUOTATION_FILENAME
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '融資餘額', '融券餘額']
	DEFAULT_CB_STOCK_QUOTATION_FIELD_TYPE = [str, float, float, int, float, float, int, int,]


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


	@classmethod
	def __is_cb(cls, cb_id):
		return True if (re.match("[\d]{5}", cb_id) is not None) else False


	@classmethod
	def __get_conversion_ratio(cls, conversion_price):
		return float(100.0 / conversion_price)


	@classmethod
	def __get_conversion_premium_rate(cls, conversion_price, bond_price, share_price):
		conversion_ratio = cls.__get_conversion_ratio(conversion_price)
		conversion_value = float(conversion_ratio * share_price)
		conversion_premium_rate = float(bond_price - conversion_value) / conversion_value
		return conversion_premium_rate


	def __init__(self, cfg):
		self.xcfg = {
			"cb_folderpath": None,
			"cb_summary_filename": None,
			"cb_publish_filename": None,
			"cb_quotation_filename": None,
			"cb_stock_quotation_filename": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["cb_folderpath"] = self.DEFAULT_CB_FOLDERPATH if self.xcfg["cb_folderpath"] is None else self.xcfg["cb_folderpath"]
		self.xcfg["cb_summary_filename"] = self.DEFAULT_CB_SUMMARY_FULL_FILENAME if self.xcfg["cb_summary_filename"] is None else self.xcfg["cb_summary_filename"]
		self.xcfg["cb_summary_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_summary_filename"])
		self.xcfg["cb_publish_filename"] = self.DEFAULT_CB_PUBLISH_FULL_FILENAME if self.xcfg["cb_publish_filename"] is None else self.xcfg["cb_publish_filename"]
		self.xcfg["cb_publish_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_publish_filename"])
		self.xcfg["cb_quotation_filename"] = self.DEFAULT_CB_QUOTATION_FULL_FILENAME if self.xcfg["cb_quotation_filename"] is None else self.xcfg["cb_quotation_filename"]
		self.xcfg["cb_quotation_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_quotation_filename"])
		self.xcfg["cb_stock_quotation_filename"] = self.DEFAULT_CB_STOCK_QUOTATION_FULL_FILENAME if self.xcfg["cb_stock_quotation_filename"] is None else self.xcfg["cb_stock_quotation_filename"]
		self.xcfg["cb_stock_quotation_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_stock_quotation_filename"])

		file_not_exist_list = []
		if not self. __check_file_exist(self.xcfg["cb_summary_filepath"]):
			file_not_exist_list.append(self.xcfg["cb_summary_filepath"])
		# else:
		# 	print ("Read CB Sumary from: %s" % self.xcfg["cb_summary_filepath"])
		if not self. __check_file_exist(self.xcfg["cb_publish_filepath"]):
			file_not_exist_list.append(self.xcfg["cb_publish_filepath"])
		# else:
		# 	print ("Read CB Publish from: %s" % self.xcfg["cb_publish_filepath"])
		if not self. __check_file_exist(self.xcfg["cb_quotation_filepath"]):
			file_not_exist_list.append(self.xcfg["cb_quotation_filepath"])
		# else:
		# 	print ("Read CB Quotation from: %s" % self.xcfg["cb_quotation_filepath"])
		if not self. __check_file_exist(self.xcfg["cb_stock_quotation_filepath"]):
			file_not_exist_list.append(self.xcfg["cb_stock_quotation_filepath"])
		# else:
		# 	print ("Read CB Stcok Quotation from: %s" % self.xcfg["cb_stock_quotation_filepath"])
		# import pdb; pdb.set_trace()
		if len(file_not_exist_list) > 0:
			raise RuntimeError("The file[%s] does NOT exist" % ", ".join(file_not_exist_list))

		self.cb_summary = self.__read_cb_summary()
		self.cb_publish = self.__read_cb_publish()
		self.cb_workbook = None
		self.cb_worksheet = None
		self.cb_stock_workbook = None
		self.cb_stock_worksheet = None
		self.cb_id_list = None  # list(self.cb_summary.keys())
		self.cb_stock_id_list = None  


	def __enter__(self):
# Open the workbook of cb quotation
		self.cb_workbook = xlrd.open_workbook(self.xcfg["cb_quotation_filepath"])
		self.cb_worksheet = self.cb_workbook.sheet_by_index(0)
# Open the workbook of cb stock quotation
		self.cb_stock_workbook = xlrd.open_workbook(self.xcfg["cb_stock_quotation_filepath"])
		self.cb_stock_worksheet = self.cb_stock_workbook.sheet_by_index(0)
		return self


	def __exit__(self, type, msg, traceback):
		if self.cb_workbook is not None:
			self.cb_workbook.release_resources()
			del self.cb_workbook
			self.cb_workbook = None
		if self.cb_stock_workbook is not None:
			self.cb_stock_workbook.release_resources()
			del self.cb_stock_workbook
			self.cb_stock_workbook = None
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
# ['可轉債商品', '到期日', '可轉換日', '票面利率', '上次付息日', '轉換價格', '現股收盤價', '可轉債價格', '套利報酬', '年化殖利率', '']
					# print(title_list)
				else:
					assert title_list is not None, "title_list should NOT be None"
					data_list = list(map(lambda x: x.lstrip("=\"").rstrip("\""), row))
					for data_index, data_value in enumerate(data_list):
						try:
							data_type = self.DEFAULT_CB_SUMMARY_FIELD_TYPE[data_index]
							data_value = data_type(data_value)
						except ValueError:
							# print "End row index: %d" % row_index
							data_value = None
						data_list[data_index] = data_value
					mobj = re.match(regex, data_list[0])
					if mobj is None: 
						raise ValueError("Incorrect format: %s" % data_list[0])
					data_list[0] = mobj.group(1)
					data_dict = dict(zip(title_list, data_list))
					cb_data[mobj.group(2)] = data_dict
				# print ("%s" % (",".join(data_list)))
		return cb_data


	def __read_cb_publish(self):
		pattern = "([\d]+)年"
		cb_data = {}
		with open(self.xcfg["cb_publish_filepath"], newline='') as f:
			rows = csv.reader(f)
			regex = re.compile(pattern)
			title_list = None
			title_tenor_index = None
			title_par_value_index = None
			for index, row in enumerate(rows):
				if index in [0, 1, 3,]: pass
				elif index == 2:
					title_list = row
					title_list = title_list[1:]  # ignore 債券代號
					title_tenor_index = title_list.index("年期")
					title_par_value_index = title_list.index("發行總面額")
# ['債券簡稱', '發行人', '發行日期', '到期日期', '年期', '發行總面額', '發行資料']
					# print(title_list)
				else:
					assert title_list is not None, "title_list should NOT be None"
					data_list = []
					data_key = row[0]
					for data_index, data_value in enumerate(row[1:]):  # ignore 債券代號
						try:
							if data_index == title_tenor_index:
								mobj = re.match(regex, data_value)
								# import pdb; pdb.set_trace()
								if mobj is None: 
									raise ValueError("Incorrect format in 年期 field: %s" % data_value)
								data_value = mobj.group(1)
							elif data_index == title_par_value_index:
								data_value = data_value.replace(",","")
							data_type = self.DEFAULT_CB_PUBLISH_FIELD_TYPE[data_index]
							data_value = data_type(data_value)
							data_list.append(data_value)
						except ValueError as e:
							print ("Exception occurs in %s, due to: %s" % (data_key, str(e)))
							raise e						
					data_dict = dict(zip(title_list, data_list))
					cb_data[data_key] = data_dict
		# import pdb; pdb.set_trace()
		return cb_data


	def __read_worksheet(self, worksheet, check_data_filter_funcptr):
		worksheet_data = {}
		# import pdb; pdb.set_trace()
		title_list = []
# title
		for column_index in range(1, worksheet.ncols):
			title_value = worksheet.cell_value(0, column_index)
			title_list.append(title_value)
		# print(title_list)
		# import pdb; pdb.set_trace()
# data
		for row_index in range(1, worksheet.nrows):
			data_key = worksheet.cell_value(row_index, 0)
			data_list = []
			for column_index in range(1, worksheet.ncols):
				data_value = worksheet.cell_value(row_index, column_index)
				data_list.append(data_value)
			# print ("%s: %s" % (data_key, data_list))
			try:
				check_data_filter_funcptr(data_list)
			except ValueError as e:
				print ("Exception occurs in %s" % data_key)
				raise e
			data_dict = dict(zip(title_list, data_list))
			worksheet_data[data_key] = data_dict
		return worksheet_data


	def __check_cb_quotation_data(self, data_list):
		data_index = 0
		for data_value in data_list:
			try:
				data_type = self.DEFAULT_CB_QUOTATION_FIELD_TYPE[data_index]
				data_value = data_type(data_value)
				data_list[data_index] = data_value
			except ValueError:
				data_list[data_index] = None
			except Exception as e:
				# import pdb; pdb.set_trace()
				traceback.print_exc()
				raise e
			data_index += 1


	def __read_cb_quotation(self):
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '到期日']
		cb_data_dict = self.__read_worksheet(self.cb_worksheet, self.__check_cb_quotation_data)
		if self.cb_id_list is None:
			cb_id_list = list(cb_data_dict.keys())
			self.cb_id_list = list(filter(self.__is_cb, cb_id_list))
		return cb_data_dict


	def __check_cb_stock_quotation_data(self, data_list):
		data_index = 0
		for data_value in data_list:
			try:
				data_type = self.DEFAULT_CB_STOCK_QUOTATION_FIELD_TYPE[data_index]
				data_value = data_type(data_value)
				data_list[data_index] = data_value
			except ValueError as e:
					# print "End row index: %d" % row_index
				if data_index == 4:  # 買進一
					# print(data_list)
					# import pdb; pdb.set_trace()
					if re.match("市價", data_value) is not None:  # 漲停
						data_list[4] = data_list[1]  # 買進一 設為 成交價
						data_list[5] = None
						break
					else:
						if re.match("市價", data_list[5]) is None:  # 不是跌停
							# import pdb; pdb.set_trace()
							traceback.print_exc()
							raise e
						else: 
							data_list[4] = None
				elif data_index == 5:  # 賣出一
					# print(data_list)
					# import pdb; pdb.set_trace()
					if re.match("市價", data_value) is not None:  # 跌停
						data_list[5] = data_list[1]  # 賣出一 設為 成交價
						assert data_list[4] == None, "買進一 should be None"
					else:
						traceback.print_exc()
						raise e
				else:
					data_list[data_index] = None
			except Exception as e:
				# import pdb; pdb.set_trace()
				traceback.print_exc()
				raise e
			data_index += 1


	def __read_cb_stock_quotation(self):
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '融資餘額', '融券餘額']
		cb_stock_data_dict = self.__read_worksheet(self.cb_stock_worksheet, self.__check_cb_stock_quotation_data)
		if self.cb_stock_id_list is None:
			self.cb_stock_id_list = list(cb_stock_data_dict.keys())
		return cb_stock_data_dict


	def check_cb_quotation_table_field(self, cb_quotation_data):
		# import pdb; pdb.set_trace()
		cb_summary_id_set = set(self.cb_summary.keys())
		cb_publish_id_set = set(self.cb_publish.keys())
		cb_quotation_id_set = set(self.cb_id_list)
		cb_diff_id_set1 = cb_summary_id_set - cb_quotation_id_set
		if len(cb_diff_id_set1) > 0:
			# raise ValueError("The CB keys are NOT identical: %s" % cb_diff_id_set)
			print("The CB IDs are NOT identical[1]: %s" % cb_diff_id_set1)
			# for cb_id in list(cb_diff_id_set):
			# 	self.cb_id_list.remove(cb_id)
		cb_diff_id_set2 = cb_publish_id_set - cb_quotation_id_set
		if len(cb_diff_id_set2) > 0:
			# raise ValueError("The CB keys are NOT identical: %s" % cb_diff_id_set)
			print("The CB IDs are NOT identical[2]: %s" % cb_diff_id_set2)


	def check_cb_stock_quotation_table_field(self, cb_quotation_data, cb_stock_quotation_data):
		# import pdb; pdb.set_trace()
		new_stock_id_set = set(map(lambda x: x[:4], self.cb_id_list))
		old_stock_id_set = set(self.cb_stock_id_list)
		deleted_stock_id_set = old_stock_id_set - new_stock_id_set
		added_stock_id_set = new_stock_id_set - old_stock_id_set
		stock_changed = False
		if len(deleted_stock_id_set) > 0:
			print("The stocks are deleted: %s" % deleted_stock_id_set)
			if not stock_changed: stock_changed = True
		if len(added_stock_id_set) > 0:
			print("The stocks are added: %s" % added_stock_id_set)
			if not stock_changed: stock_changed = True
		if stock_changed:
			raise ValueError("The stocks are NOT identical")


	def calculate_internal_rate_of_return(self, cb_quotation, use_percentage=True):
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
			irr_dict[cb_id] = {"商品": cb_quotation_data["商品"], "到期日": cb_quotation_data["到期日"], "賣出一": cb_quotation_data["賣出一"], "到期天數": days, "年化報酬率": irr}
		return irr_dict


	def get_positive_internal_rate_of_return(self, cb_quotation, positive_threshold=1, duration_within_days=365, need_sort=True):
		irr_dict = self.calculate_internal_rate_of_return(cb_quotation, use_percentage=True)
		if positive_threshold is not None:
			irr_dict = dict(filter(lambda x: x[1]["年化報酬率"] > positive_threshold, irr_dict.items()))
		if duration_within_days is not None:
			# duration_date = datetime.now() + timedelta(days=duration_within_days)
			# irr_dict = dict(filter(lambda x: datetime.strptime(x[1]["到期日"],"%Y/%m/%d") <= duration_date, irr_dict.items()))
			irr_dict = dict(filter(lambda x: x[1]["到期天數"] <= duration_within_days, irr_dict.items()))
		if need_sort:
			irr_dict = collections.OrderedDict(sorted(irr_dict.items(), key=lambda x: x[1]["年化報酬率"], reverse=True))
		return irr_dict


	def calculate_premium(self, cb_quotation, cb_stock_quotation, need_breakeven=False, use_percentage=True):
		premium_dict = {}
		# import pdb; pdb.set_trace()
		for cb_id in self.cb_id_list:
			cb_quotation_data = cb_quotation[cb_id]
			cb_summary_data = self.cb_summary[cb_id]
			cb_stock_id = cb_id[:4]
			cb_stock_quotation_data = cb_stock_quotation[cb_stock_id]
			if cb_stock_quotation_data["買進一"] is None:
				# print("Ignore CB Stock[%s]: 沒有 買進一" % cb_stock_id)
				continue
			if cb_quotation_data["賣出一"] is None:
				# print("Ignore CB[%s]: 沒有 賣出一" % cb_id)
				continue
			if need_breakeven and cb_quotation_data["成交"] is None:
				# print("Ignore CB[%s]: 沒有 成交" % cb_stock_id)
				continue
			# conversion_parity = self.__get_conversion_parity(cb_summary_data["轉換價格"], cb_stock_quotation_data["買進一"])
			# premium = (cb_quotation_data["賣出一"] - conversion_parity) / conversion_parity
			conversion_premium_rate = self.__get_conversion_premium_rate(cb_summary_data["轉換價格"], cb_quotation_data["賣出一"], cb_stock_quotation_data["買進一"])
			days = self.__get_days(cb_quotation_data["到期日"])
			if use_percentage:
				conversion_premium_rate *= 100.0
			premium_dict[cb_id] = {
				"商品": cb_quotation_data["商品"], 
				"到期日": cb_quotation_data["到期日"], 
				"到期天數": days, 
				"溢價率": conversion_premium_rate, 
				# "成交": cb_quotation_data["成交"], 
				"賣出一": cb_quotation_data["賣出一"], 
				"融資餘額": cb_stock_quotation_data["融資餘額"], 
				"融券餘額": cb_stock_quotation_data["融券餘額"]
			}
			if need_breakeven: premium_dict[cb_id]["成交"] = cb_quotation_data["成交"]
		return premium_dict


	def get_negative_premium(self, cb_quotation, cb_stock_quotation, negative_threshold=-1, need_sort=True):
		premium_dict = self.calculate_premium(cb_quotation, cb_stock_quotation, use_percentage=True)
		premium_dict = dict(filter(lambda x: x[1]["溢價率"] <= negative_threshold, premium_dict.items()))
		if need_sort:
			premium_dict = collections.OrderedDict(sorted(premium_dict.items(), key=lambda x: x[1]["溢價率"], reverse=False))
		return premium_dict


	def get_low_premium_and_breakeven(self, cb_quotation, cb_stock_quotation, low_threshold=5, breakeven_threshold=105, need_sort=True):
		premium_dict = self.calculate_premium(cb_quotation, cb_stock_quotation, need_breakeven=True, use_percentage=True)
		premium_dict = dict(filter(lambda x: x[1]["溢價率"] <= low_threshold and x[1]["成交"] <= breakeven_threshold, premium_dict.items()))
		if need_sort:
			premium_dict = collections.OrderedDict(sorted(premium_dict.items(), key=lambda x: x[1]["溢價率"], reverse=False))
		return premium_dict


	def calculate_stock_premium(self, cb_quotation, cb_stock_quotation, use_percentage=True):
		stock_premium_dict = {}
		# import pdb; pdb.set_trace()
		for cb_id in self.cb_id_list:
			cb_quotation_data = cb_quotation[cb_id]
			cb_summary_data = self.cb_summary[cb_id]
			cb_stock_id = cb_id[:4]
			cb_stock_quotation_data = cb_stock_quotation[cb_stock_id]
			if cb_stock_quotation_data["成交"] is None:
				# print("Ignore CB Stock[%s]: 沒有 成交" % cb_stock_id)
				continue
			days = self.__get_days(cb_quotation_data["到期日"])
			stock_premium = (cb_stock_quotation_data["成交"] - cb_summary_data["轉換價格"]) / cb_summary_data["轉換價格"]
			# print(cb_quotation_data)
			if use_percentage:
				stock_premium *= 100.0
			stock_premium_dict[cb_id] = {"商品": cb_quotation_data["商品"], "到期日": cb_quotation_data["到期日"], "到期天數": days, "股票溢價率": stock_premium}
		return stock_premium_dict


	def get_absolute_stock_premium(self, cb_quotation, cb_stock_quotation, absolute_threshold=5, duration_within_days=180, need_sort=True):
		stock_premium_dict = self.calculate_stock_premium(cb_quotation, cb_stock_quotation, use_percentage=True)
		stock_premium_dict = dict(filter(lambda x: abs(x[1]["股票溢價率"]) <= absolute_threshold, stock_premium_dict.items()))
		if duration_within_days is not None:
			# duration_date = datetime.now() + timedelta(days=duration_within_days)
			# irr_dict = dict(filter(lambda x: datetime.strptime(x[1]["到期日"],"%Y/%m/%d") <= duration_date, irr_dict.items()))
			stock_premium_dict = dict(filter(lambda x: x[1]["到期天數"] <= duration_within_days, stock_premium_dict.items()))
		if need_sort:
			stock_premium_dict = collections.OrderedDict(sorted(stock_premium_dict.items(), key=lambda x: x[1]["股票溢價率"], reverse=False))
		return stock_premium_dict


	def check_data_source(self, cb_quotation_data, cb_stock_quotation_data):
		self.check_cb_quotation_table_field(cb_quotation_data)
		self.check_cb_stock_quotation_table_field(cb_quotation_data, cb_stock_quotation_data)


	def test(self):
		# data_dict_summary = self.__read_cb_summary()
		# # print (data_dict_summary)
		data_dict_quotation = self.__read_cb_quotation()
		stock_data_dict_quotation = self.__read_cb_stock_quotation()
		self.check_data_source(data_dict_quotation, stock_data_dict_quotation)
		# print (data_dict_quotation)
		# print(self.calculate_internal_rate_of_return(data_dict_quotation))
		print("=== 年化報酬率 ==================================================")
		irr_dict = self.get_positive_internal_rate_of_return(data_dict_quotation)
		for irr_key, irr_data in irr_dict.items():
			print ("%s[%s]: %.2f  %.2f  %s" % (irr_data["商品"], irr_key, float(irr_data["年化報酬率"]), float(irr_data["賣出一"]), irr_data["到期日"]))
		print("=================================================================\n")
		print("=== 溢價率(套利) ======================================================")
		premium_dict = self.get_negative_premium(data_dict_quotation, stock_data_dict_quotation)
		for premium_key, premium_data in premium_dict.items():
			print ("%s[%s]: %.2f  %d  %d" % (premium_data["商品"], premium_key, float(premium_data["溢價率"]), premium_data["融資餘額"], premium_data["融券餘額"]))
		print("=================================================================\n")
		print("=== 股票溢價率 ==================================================")
		stock_premium_dict = self.get_absolute_stock_premium(data_dict_quotation, stock_data_dict_quotation)
		for stock_premium_key, stock_premium_data in stock_premium_dict.items():
			print ("%s[%s]: %.2f" % (stock_premium_data["商品"], stock_premium_key, float(stock_premium_data["股票溢價率"])))
		print("=================================================================\n")
		print("=== 低溢價且保本 ================================================")
		cb_dict = self.get_low_premium_and_breakeven(data_dict_quotation, stock_data_dict_quotation)
		for cb_key, cb_data in cb_dict.items():
			print ("%s[%s]: %.2f  %.2f  %.2f  %s" % (cb_data["商品"], cb_key, float(cb_data["溢價率"]), float(cb_data["成交"]), float(cb_data["賣出一"]), cb_data["到期日"]))
		print("=================================================================\n")


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
	# parser.add_argument('--list_analysis_method', required=False, action='store_true', help='List each analysis method and exit')
	# args = parser.parse_args()

	cfg = {
	}
	with ConvertibleBondAnalysis(cfg) as obj:
		obj.test()

	# from selenium import webdriver
	# import time

	# driver = webdriver.Chrome("C:\chromedriver.exe")
	# # driver.get("https://www.tpex.org.tw/web/bond/publish/convertible_bond_search/memo.php?l=zh-tw")
	# driver.get("https://mops.twse.com.tw/mops/web/t120sg01?TYPEK=&bond_id=45552&bond_kind=5&bond_subn=%24M00000001&bond_yrn=2&come=2&encodeURIComponent=1&firstin=ture&issuer_stock_code=4555&monyr_reg=202302&pg=&step=0&tg=k_code=4555&monyr_reg=202302&pg=&step=0&tg=")
	# time.sleep(5)
	# # #找到輸入框
	# # element = driver.find_element_by_name("q");
	# # #輸入內容
	# # element.send_keys("hello world");
	# # #提交表單
	# # element.submit();
	# driver.close()

