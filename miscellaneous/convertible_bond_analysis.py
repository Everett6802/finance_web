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
import time
import json
from collections import OrderedDict


class ConvertibleBondAnalysis(object):

	DEFAULT_CONFIG_FOLDERPATH =  "C:\\Users\\%s" % os.getlogin()
	DEFAULT_DISPLAY_CB_LIST_FILENAME = "convertible_bond_list.txt"

	DEFAULT_CB_FOLDERPATH =  "C:\\可轉債"
	DEFAULT_CB_DATA_FOLDERNAME =  "Data"
	DEFAULT_CB_SUMMARY_FILENAME = "可轉債總表"	
	DEFAULT_CB_SUMMARY_FULL_FILENAME = "%s.csv" % DEFAULT_CB_SUMMARY_FILENAME
	DEFAULT_CB_MONTHLY_CONVERT_DATA_FILENAME_PREFIX = "可轉換公司債月分析表"

# ['可轉債商品', '到期日', '可轉換日', '票面利率', '上次付息日', '轉換價格', '現股收盤價', '可轉債價格', '套利報酬', '年化殖利率', '']
	DEFAULT_CB_SUMMARY_FIELD_TYPE = [str, str, str, float, str, float, float, float, float, float, str,]
	DEFAULT_CB_SUMMARY_FIELD_TYPE_LEN = len(DEFAULT_CB_SUMMARY_FIELD_TYPE)
	DEFAULT_CB_PUBLISH_FILENAME = "可轉債發行"
	DEFAULT_CB_PUBLISH_FULL_FILENAME = "%s.csv" % DEFAULT_CB_PUBLISH_FILENAME
# ['債券簡稱', '發行人', '發行日期', '到期日期', '年期', '發行總面額', '發行資料']
	DEFAULT_CB_PUBLISH_FIELD_TYPE = [str, str, str, str, int, int, str,]
	DEFAULT_CB_PUBLISH_FIELD_TYPE_LEN = len(DEFAULT_CB_PUBLISH_FIELD_TYPE)
	DEFAULT_CB_QUOTATION_FILENAME = "可轉債報價"
	DEFAULT_CB_QUOTATION_FULL_FILENAME = "%s.xlsx" % DEFAULT_CB_QUOTATION_FILENAME
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '到期日']
	DEFAULT_CB_QUOTATION_FIELD_TYPE = [str, float, float, int, float, float, str,]
	DEFAULT_CB_QUOTATION_FIELD_TYPE_LEN = len(DEFAULT_CB_QUOTATION_FIELD_TYPE)
	DEFAULT_CB_STOCK_QUOTATION_FILENAME = "可轉債個股報價"
	DEFAULT_CB_STOCK_QUOTATION_FULL_FILENAME = "%s.xlsx" % DEFAULT_CB_STOCK_QUOTATION_FILENAME
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '融資餘額', '融券餘額']
	DEFAULT_CB_STOCK_QUOTATION_FIELD_TYPE = [str, float, float, int, float, float, int, int,]
	DEFAULT_CB_STOCK_QUOTATION_FIELD_TYPE_LEN = len(DEFAULT_CB_STOCK_QUOTATION_FIELD_TYPE)

	CB_PUBLISH_DETAIL_URL_FORMAT = "https://mops.twse.com.tw/mops/web/t120sg01?TYPEK=&bond_id=%s&bond_kind=5&bond_subn=%24M00000001&bond_yrn=5&come=2&encodeURIComponent=1&firstin=ture&issuer_stock_code=%s&monyr_reg=%s&pg=&step=0&tg="
	CB_TRADING_SUSPENSION_SET = {"30184"}

	STATEMENT_RELEASE_DATE_LIST = [(3,31,),(5,15,),(8,14,),(11,14),]

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
	def __get_days(cls, end_date_str=None, start_date_str=None):
		# import pdb; pdb.set_trace()
		end_date = datetime.strptime(end_date_str,"%Y/%m/%d") if end_date_str is not None else datetime.fromordinal(date.today().toordinal())
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


	@classmethod
	def __check_request_module_installed(cls):
		try:
			module = __import__("requests")
		except ModuleNotFoundError:
			return False
		return True


	@classmethod
	def __check_bs4_module_installed(cls):
		try:
			module = __import__("bs4")
		except ModuleNotFoundError:
			return False
		return True


	@classmethod
	def __can_scrape(cls):
		if not cls.__check_request_module_installed(): return False
		if not cls.__check_bs4_module_installed(): return False
		return True


	@classmethod
	def __check_selenium_module_installed(cls):
		try:
			module = __import__("selenium.webdriver")
		except ModuleNotFoundError:
			return False
		return True


	@classmethod
	def __can_scrape_ui(cls):
		if not cls.__check_selenium_module_installed(): return False
		return True


	@classmethod
	def __check_wget_module_installed(cls):
		try:
			module = __import__("wget")
		except ModuleNotFoundError:
			return False
		return True


	# @classmethod
	# def __get_web_driver(cls, web_driver_filepath="C:\chromedriver.exe"):
	# 	module = __import__("selenium.webdriver")
	# 	web_driver_class = getattr(module, "webdriver")
	# 	web_driver_obj = web_driver_class.Chrome(web_driver_filepath)
	# 	return web_driver_obj


	def __init__(self, cfg):
		self.xcfg = {
			"cb_folderpath": None,
			"cb_data_folderpath": None,
			"cb_summary_filename": None,
			"cb_publish_filename": None,
			"cb_quotation_filename": None,
			"cb_stock_quotation_filename": None,
			"cb_all": False,
			"cb_list_filename": None,
			"cb_list": None
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["cb_folderpath"] = self.DEFAULT_CB_FOLDERPATH if self.xcfg["cb_folderpath"] is None else self.xcfg["cb_folderpath"]
		self.xcfg["cb_data_folderpath"] = os.path.join(self.xcfg["cb_folderpath"], self.DEFAULT_CB_DATA_FOLDERNAME) if self.xcfg["cb_data_folderpath"] is None else self.xcfg["cb_data_folderpath"]
		self.xcfg["cb_summary_filename"] = self.DEFAULT_CB_SUMMARY_FULL_FILENAME if self.xcfg["cb_summary_filename"] is None else self.xcfg["cb_summary_filename"]
		self.xcfg["cb_summary_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_summary_filename"])
		self.xcfg["cb_publish_filename"] = self.DEFAULT_CB_PUBLISH_FULL_FILENAME if self.xcfg["cb_publish_filename"] is None else self.xcfg["cb_publish_filename"]
		self.xcfg["cb_publish_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_publish_filename"])
		self.xcfg["cb_quotation_filename"] = self.DEFAULT_CB_QUOTATION_FULL_FILENAME if self.xcfg["cb_quotation_filename"] is None else self.xcfg["cb_quotation_filename"]
		self.xcfg["cb_quotation_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_quotation_filename"])
		self.xcfg["cb_stock_quotation_filename"] = self.DEFAULT_CB_STOCK_QUOTATION_FULL_FILENAME if self.xcfg["cb_stock_quotation_filename"] is None else self.xcfg["cb_stock_quotation_filename"]
		self.xcfg["cb_stock_quotation_filepath"] = os.path.join(self.xcfg["cb_folderpath"], self.xcfg["cb_stock_quotation_filename"])
		self.xcfg["cb_list_filename"] = self.DEFAULT_DISPLAY_CB_LIST_FILENAME if self.xcfg["cb_list_filename"] is None else self.xcfg["cb_list_filename"]
		self.xcfg["cb_list_filepath"] = os.path.join(self.DEFAULT_CONFIG_FOLDERPATH, self.xcfg["cb_list_filename"])
# Check file exist
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

		if not os.path.exists(self.xcfg["cb_data_folderpath"]):
			print("The CB data folder[%s] does NOT exist" % self.xcfg["cb_data_folderpath"])
			os.mkdir(self.xcfg["cb_data_folderpath"])

		self.cb_summary = self.__read_cb_summary()
		self.cb_publish = self.__read_cb_publish()
		self.cb_workbook = None
		self.cb_worksheet = None
		self.cb_stock_workbook = None
		self.cb_stock_worksheet = None
		self.cb_id_list = None  # list(self.cb_summary.keys())
		self.cb_stock_id_list = None

		self.can_scrape = self.__can_scrape()
		self.requests_module = None
		self.beautifulsoup_class = None
		self.web_driver = None
		self.selenium_select_module = None
		self.wget_module = None
		self.cb_publish_detail = {}

		self.STOCK_INFO_SCRAPY_METADATA_DICT = {
# Daily
			"法人持股": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcl/zcl.djhtm?a=%s&b=3",
				"SCRAPY_FUNCPTR": self.__stock_info_cooperate_shareholding_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Daily",
			},
			"主力進出": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zco/zco_%s.djhtm",
				"SCRAPY_FUNCPTR": self.__stock_info_major_inflow_outflow_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Daily",
			},
			# "融資融券": "https://concords.moneydj.com/z/zc/zcx/zcx_%s.djhtm",
			# "融資融券": "https://concords.moneydj.com/z/zc/zcn/zcn_%s.djhtm",
			"融資融券": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcn/zcn.djhtm?a=%s&b=3",
				"SCRAPY_FUNCPTR": self.__stock_info_margin_trading_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Daily",
			},
# Monthly
			"月營收": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zch/zch_%s.djhtm",
				"SCRAPY_FUNCPTR": self.__stock_info_revenue_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Monthly",
			},
# Quarterly
			"獲利能力": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zce/zce_%s.djhtm",
				"SCRAPY_FUNCPTR": self.__stock_info_profitability_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Quarterly",
			},
			"季盈餘": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zch/zch_%s.djhtm",
				"SCRAPY_FUNCPTR": self.__stock_info_earning_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Quarterly",
			},
			"資產負債簡表(季)": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcp/zcp.djhtm?a=%s&b=1&c=Q",
				"SCRAPY_FUNCPTR": self.__stock_info_balance_sheet_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Quarterly",
			},
			"現金流量簡表(季)": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcp/zcp.djhtm?a=%s&b=3&c=Q",
				"SCRAPY_FUNCPTR": self.__stock_info_cash_flow_statement_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Quarterly",
			},
			"財務比率簡表(季)": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcp/zcp0.djhtm?a=%s&c=Q",
				"SCRAPY_FUNCPTR": self.__stock_info_financial_ratio_statement_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Quarterly",
			},
# Yearly
			"資產負債簡表(年)": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcp/zcp.djhtm?a=%s&b=1&c=Y",
				"SCRAPY_FUNCPTR": self.__stock_info_balance_sheet_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Yearly",
			},
			"現金流量簡表(年)": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcp/zcp.djhtm?a=%s&b=3&c=Y",
				"SCRAPY_FUNCPTR": self.__stock_info_cash_flow_statement_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Yearly",
			},
			"財務比率簡表(年)": {
				"URL_FORMAT": "https://concords.moneydj.com/z/zc/zcp/zcp0.djhtm?a=%s&c=Y",
				"SCRAPY_FUNCPTR": self.__stock_info_financial_ratio_statement_scrapy_funcptr,
				"UPDATE_FREQUENCY": "Yearly",
			},
		}

		self.STOCK_INFO_SCRAPY_URL_FORMAT_DICT = {key: value["URL_FORMAT"] for key, value in self.STOCK_INFO_SCRAPY_METADATA_DICT.items()}
		self.STOCK_INFO_SCRAPY_FUNCPTR_DICT = {key: value["SCRAPY_FUNCPTR"] for key, value in self.STOCK_INFO_SCRAPY_METADATA_DICT.items()}
		self.STOCK_INFO_SCRAPY_UPDATE_FREQUENCY_DICT = {key: value["UPDATE_FREQUENCY"] for key, value in self.STOCK_INFO_SCRAPY_METADATA_DICT.items()}

		self.STOCK_INFO_UPDATE_TIME_FUNCPTR_DICT = {
			"Daily": self.__calculate_stock_info_daily_update_time,
			"Monthly": self.__calculate_stock_info_monthly_update_time,
			"Quarterly": self.__calculate_stock_info_quarterly_update_time,
			"Yearly": self.__calculate_stock_info_yearly_update_time,
		}
		self.filepath_dict = OrderedDict()
		self.filepath_dict["cb_summary_filepath"] = self.xcfg["cb_summary_filepath"]
		self.filepath_dict["cb_publish_filepath"] = self.xcfg["cb_publish_filepath"]
		self.filepath_dict["cb_quotation_filepath"] = self.xcfg["cb_quotation_filepath"]
		self.filepath_dict["cb_stock_quotation_filepath"] = self.xcfg["cb_stock_quotation_filepath"]
		self.filepath_dict["cb_list_filepath"] = self.xcfg["cb_list_filepath"]
		self.filepath_dict["cb_data_filepath"] = self.xcfg["cb_data_folderpath"]
# Update the CB ID list
		if not self.xcfg['cb_all']:
			if self.xcfg["cb_list"] is not None:
				if type(self.xcfg["cb_list"]) is str:
					cb_list = []
					for cb in self.xcfg["cb_list"].split(","):
						cb_list.append(cb)
					self.xcfg["cb_list"] = cb_list
			else:	
				self.__get_cb_list_from_file()
			self.cb_id_list = self.xcfg["cb_list"]
# Check if the incorrect CB IDs exist
			illegal_cb_id_list = list(filter(lambda x: re.match("[\d]{5}", x) is None, self.cb_id_list))
			assert len(illegal_cb_id_list) == 0, "Illegal CB ID list: %s" % illegal_cb_id_list
			self.cb_stock_id_list = list(set(list(map(lambda x: x[:4], self.cb_id_list))))


	def __enter__(self):
# Open the workbook of cb quotation
		self.cb_workbook = xlrd.open_workbook(self.xcfg["cb_quotation_filepath"])
		self.cb_worksheet = self.cb_workbook.sheet_by_index(0)
# Open the workbook of cb stock quotation
		self.cb_stock_workbook = xlrd.open_workbook(self.xcfg["cb_stock_quotation_filepath"])
		self.cb_stock_worksheet = self.cb_stock_workbook.sheet_by_index(0)
		return self


	def __exit__(self, type, msg, traceback):
		if self.web_driver is not None:
			self.web_driver.close()
			self.web_driver = None
		if self.cb_workbook is not None:
			self.cb_workbook.release_resources()
			del self.cb_workbook
			self.cb_workbook = None
		if self.cb_stock_workbook is not None:
			self.cb_stock_workbook.release_resources()
			del self.cb_stock_workbook
			self.cb_stock_workbook = None
		return False


	def __get_requests_module(self):
		if not self.__check_request_module_installed():
			raise RuntimeError("The requests module is NOT installed!!!")
		if self.requests_module is None:
			self.requests_module = __import__("requests")
		return self.requests_module


	def __get_beautifulsoup_class(self):
		if not self.__check_bs4_module_installed():
			raise RuntimeError("The bs4 module is NOT installed!!!")
		if self.beautifulsoup_class is None:
			bs4_module = __import__("bs4")
			self.beautifulsoup_class = getattr(bs4_module, "BeautifulSoup")
		return self.beautifulsoup_class


	def __get_web_driver(self, web_driver_filepath="C:\chromedriver.exe"):
		if not self.__check_selenium_module_installed():
			raise RuntimeError("The selenium module is NOT installed!!!")
		if self.web_driver is None:
			module = __import__("selenium.webdriver")
			web_driver_class = getattr(module, "webdriver")
			self.web_driver = web_driver_class.Chrome(web_driver_filepath)
		return self.web_driver


	def __get_selenium_select_module(self):
		if not self.__check_selenium_module_installed():
			raise RuntimeError("The selenium module is NOT installed!!!")
		# import pdb; pdb.set_trace()
		if self.selenium_select_module is None:
			self.selenium_select_module = __import__('selenium.webdriver.support.ui', fromlist=['Select'])
		return self.selenium_select_module


	def __get_wget_module(self):
		if not self.__check_wget_module_installed():
			raise RuntimeError("The wget module is NOT installed!!!")
		if self.wget_module is None:
			self.wget_module = __import__("wget")
		return self.wget_module


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
						if data_index >= self.DEFAULT_CB_SUMMARY_FIELD_TYPE_LEN: break
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
						if data_index >= self.DEFAULT_CB_PUBLISH_FIELD_TYPE_LEN: break
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
			except Exception as e:
				print ("Exception occurs in %s: %s" % (data_key, data_list))
				raise e
			data_dict = dict(zip(title_list, data_list))
			worksheet_data[data_key] = data_dict
		return worksheet_data


	def __check_cb_quotation_data(self, data_list):
		data_index = 0
		for data_value in data_list:
			if data_index >= self.DEFAULT_CB_QUOTATION_FIELD_TYPE_LEN: break
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
			assert self.xcfg['cb_all'], "Incorrect setting: CB ID list"
			cb_id_list = list(cb_data_dict.keys())
			self.cb_id_list = list(filter(self.__is_cb, cb_id_list))
		return cb_data_dict


	def __check_cb_stock_quotation_data(self, data_list):
		data_index = 0
		for data_value in data_list:
			if data_index >= self.DEFAULT_CB_STOCK_QUOTATION_FIELD_TYPE_LEN: break
			try:
				data_type = self.DEFAULT_CB_STOCK_QUOTATION_FIELD_TYPE[data_index]
				data_value = data_type(data_value)
				data_list[data_index] = data_value
			except ValueError as e:
				# import pdb; pdb.set_trace()
				if data_index == 4:  # 買進一
					# print(data_list)
					# import pdb; pdb.set_trace()
					if re.match("市價", str(data_value)) is not None:  # 漲停
						data_list[4] = data_list[1]  # 買進一 設為 成交價
						data_list[5] = None
						break
					elif re.match("--", str(data_value)) is not None and isinstance(data_list[5], float):  # 跌停
						data_list[4] = None
						break
					else:
						if re.match("市價", str(data_list[5])) is None:  # 不是跌停
							# import pdb; pdb.set_trace()
							traceback.print_exc()
							raise e
						else: 
							data_list[4] = None
					# if re.match("--", data_value) is not None:  # 跌停
					# 	if int(data_list[5] * 100) != int(data_list[1] * 100):  # 不是跌停
					# 		traceback.print_exc()
					# 		raise e
					# 	data_list[4] = None
				elif data_index == 5:  # 賣出一
					# print(data_list)
					# import pdb; pdb.set_trace()
					if re.match("市價", data_value) is not None:  # 跌停
						data_list[5] = data_list[1]  # 賣出一 設為 成交價
						assert data_list[4] == None, "買進一 should be None"
					elif re.match("--", str(data_value)) is not None and isinstance(data_list[4], float):  # 漲停
						data_list[5] = None
						break
					else:
						traceback.print_exc()
						raise e
					# if re.match("--", data_value) is not None:  # 漲停
					# 	if int(data_list[4] * 100) != int(data_list[1] * 100):  # 不是漲停
					# 		traceback.print_exc()
					# 		raise e
					# 	data_list[5] = None
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
			assert self.xcfg['cb_all'], "Incorrect setting: CB stock ID list"
			self.cb_stock_id_list = list(cb_stock_data_dict.keys())
		return cb_stock_data_dict


	def check_cb_quotation_table_field(self, cb_quotation_data):
		# import pdb; pdb.set_trace()
		cb_summary_id_set = set(self.cb_summary.keys()) - self.CB_TRADING_SUSPENSION_SET
		cb_publish_id_set = set(self.cb_publish.keys()) - self.CB_TRADING_SUSPENSION_SET
		cb_quotation_id_set = set(self.cb_id_list) - self.CB_TRADING_SUSPENSION_SET
		cb_summary_diff_quotation_id_set = cb_summary_id_set - cb_quotation_id_set
		stock_changed = False
		if len(cb_summary_diff_quotation_id_set) > 0:
			print("The CB IDs are NOT identical[1]: %s" % cb_summary_diff_quotation_id_set)
			# if not stock_changed: stock_changed = True
		cb_publish_diff_quotation_id_set = cb_publish_id_set - cb_quotation_id_set
		if len(cb_publish_diff_quotation_id_set) > 0:
			cb_publish_diff_quotation_id_list = list(cb_publish_diff_quotation_id_set)
			cb_publish_diff_quotation_id_list.sort(key=lambda x: datetime.strptime(self.cb_publish[x]['發行日期'], "%Y/%m/%d"))
			new_public_cb_str_list = ["%s[%s]: %s" % (self.cb_publish[cb_publish_id]['債券簡稱'], cb_publish_id, self.cb_publish[cb_publish_id]['發行日期']) for cb_publish_id in cb_publish_diff_quotation_id_list]
			print("=== 新發行 ===")
			for new_public_cb_str in new_public_cb_str_list: print(new_public_cb_str)
			# print("The CB IDs are NOT identical[2]: %s" % cb_publish_diff_quotation_id_set)
		cb_quotation_diff_summary_id_set = cb_quotation_id_set - cb_summary_id_set
		if len(cb_quotation_diff_summary_id_set) > 0:
			print("The CB IDs are NOT identical[3]: %s" % cb_quotation_diff_summary_id_set)
			if not stock_changed: stock_changed = True
		cb_quotation_diff_publish_id_set = cb_quotation_id_set - cb_publish_id_set
		if len(cb_quotation_diff_publish_id_set) > 0:
			print("The CB IDs are NOT identical[4]: %s" % cb_quotation_diff_publish_id_set)
			if not stock_changed: stock_changed = True
		if stock_changed:
			raise ValueError("The CBs are NOT identical")


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


	def check_data_source(self, cb_quotation_data, cb_stock_quotation_data):
		self.check_cb_quotation_table_field(cb_quotation_data)
		self.check_cb_stock_quotation_table_field(cb_quotation_data, cb_stock_quotation_data)


	def calculate_internal_rate_of_return(self, cb_quotation, use_percentage=True):
		irr_dict = {}
		# import pdb; pdb.set_trace()
		for cb_id in self.cb_id_list:
			cb_quotation_data = cb_quotation[cb_id]
			# print(cb_quotation_data)
			if cb_quotation_data["賣出一"] is None:
				irr_dict[cb_id] = {"商品": cb_quotation_data["商品"], "到期日": cb_quotation_data["到期日"], "賣出一": cb_quotation_data["賣出一"], "到期天數": None, "年化報酬率": None}
				# print("%s %s" % (cb_id, str(irr_dict)))
			else:
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
			# irr_dict = dict(filter(lambda x: x[1]["年化報酬率"] is not None, irr_dict.items()))
			# irr_dict = dict(filter(lambda x: x[1]["年化報酬率"] > positive_threshold, irr_dict.items()))
			def filter_func(x):
				return True if ((x[1]["年化報酬率"] is not None) and (x[1]["年化報酬率"] > positive_threshold)) else False
			irr_dict = dict(filter(filter_func, irr_dict.items()))
		if duration_within_days is not None:
			# # duration_date = datetime.now() + timedelta(days=duration_within_days)
			# # irr_dict = dict(filter(lambda x: datetime.strptime(x[1]["到期日"],"%Y/%m/%d") <= duration_date, irr_dict.items()))
			# irr_dict = dict(filter(lambda x: x[1]["到期天數"] is not None, irr_dict.items()))
			# irr_dict = dict(filter(lambda x: x[1]["到期天數"] <= duration_within_days, irr_dict.items()))
			def filter_func(x):
				return True if ((x[1]["年化報酬率"] is not None) and (x[1]["到期天數"] <= duration_within_days)) else False
			irr_dict = dict(filter(filter_func, irr_dict.items()))
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
			if cb_stock_quotation_data["買進一"] is None or \
               cb_quotation_data["賣出一"] is None or \
               (need_breakeven and cb_quotation_data["成交"] is None):
				# print("Ignore CB Stock[%s]: 沒有 買進一" % cb_stock_id)
				# print("Ignore CB[%s]: 沒有 賣出一" % cb_id)
				# print("Ignore CB[%s]: 沒有 成交" % cb_stock_id)
				premium_dict[cb_id] = {
					"商品": cb_quotation_data["商品"], 
					"到期日": cb_quotation_data["到期日"], 
					"到期天數": None, 
					"溢價率": None, 
					# "成交": cb_quotation_data["成交"], 
					"賣出一": cb_quotation_data["賣出一"], 
					"融資餘額": cb_stock_quotation_data["融資餘額"], 
					"融券餘額": cb_stock_quotation_data["融券餘額"]
				}
			else:
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
		# premium_dict = dict(filter(lambda x: x[1]["溢價率"] is not None, premium_dict.items()))
		# premium_dict = dict(filter(lambda x: x[1]["溢價率"] <= negative_threshold, premium_dict.items()))
		def filter_func(x):
			return True if ((x[1]["溢價率"] is not None) and (x[1]["溢價率"] <= negative_threshold)) else False
		premium_dict = dict(filter(filter_func, premium_dict.items()))
		if need_sort:
			premium_dict = collections.OrderedDict(sorted(premium_dict.items(), key=lambda x: x[1]["溢價率"], reverse=False))
		return premium_dict


	def get_low_premium_and_breakeven(self, cb_quotation, cb_stock_quotation, low_conversion_premium_rate_threshold=8, breakeven_threshold=108, need_sort=True):
		premium_dict = self.calculate_premium(cb_quotation, cb_stock_quotation, need_breakeven=True, use_percentage=True)
		# premium_dict = dict(filter(lambda x: x[1]["溢價率"] is not None, premium_dict.items()))
		# premium_dict = dict(filter(lambda x: x[1]["溢價率"] <= low_conversion_premium_rate_threshold and x[1]["成交"] <= breakeven_threshold, premium_dict.items()))
		def filter_func(x):
			check = False
			if (x[1]["溢價率"] is not None) and \
			   (x[1]["成交"] is not None) and \
			   (x[1]["溢價率"] <= low_conversion_premium_rate_threshold) and \
			   (x[1]["成交"] <= breakeven_threshold):
			   check = True
			return check
		premium_dict = dict(filter(filter_func, premium_dict.items()))
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
			days = self.__get_days(cb_quotation_data["到期日"])
			stock_premium = None
			if cb_stock_quotation_data["成交"] is not None:
				# print("Ignore CB Stock[%s]: 沒有 成交" % cb_stock_id)
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
			# # duration_date = datetime.now() + timedelta(days=duration_within_days)
			# # irr_dict = dict(filter(lambda x: datetime.strptime(x[1]["到期日"],"%Y/%m/%d") <= duration_date, irr_dict.items()))
			# stock_premium_dict = dict(filter(lambda x: x[1]["到期天數"] <= duration_within_days, stock_premium_dict.items()))
			def filter_func(x):
				return True if ((x[1]["到期天數"] is not None) and (x[1]["到期天數"] <= duration_within_days)) else False
			stock_premium_dict = dict(filter(filter_func, stock_premium_dict.items()))
		if need_sort:
			stock_premium_dict = collections.OrderedDict(sorted(stock_premium_dict.items(), key=lambda x: x[1]["股票溢價率"], reverse=False))
		return stock_premium_dict


	def search_cb_opportunity_dates(self, cb_quotation, cb_stock_quotation, issuing_date_threshold=30, convertible_date_threshold=30, maturity_date_threshold=180, low_conversion_premium_rate_threshold=10, breakeven_threshold=110, use_percentage=True):
		# premium_dict = self.calculate_premium(cb_quotation, cb_stock_quotation, need_breakeven=True, use_percentage=True)
		issuing_date_cb_dict = {}
		convertible_date_cb_dict = {}
		maturity_date_cb_dict = {}
		# import pdb; pdb.set_trace()
		for cb_id in self.cb_id_list:
			cb_summary_data = self.cb_summary[cb_id]
			cb_publish_data = self.cb_publish[cb_id]
			cb_quotation_data = cb_quotation[cb_id]
			cb_stock_id = cb_id[:4]
			cb_stock_quotation_data = cb_stock_quotation[cb_stock_id]
			if cb_stock_quotation_data["買進一"] is None:
				# print("Ignore CB Stock[%s]: 沒有 買進一" % cb_stock_id)
				continue
			if cb_quotation_data["賣出一"] is None:
				# print("Ignore CB[%s]: 沒有 賣出一" % cb_id)
				continue
			if cb_quotation_data["成交"] is None:
				# print("Ignore CB[%s]: 沒有 成交" % cb_stock_id)
				continue

			conversion_premium_rate = self.__get_conversion_premium_rate(cb_summary_data["轉換價格"], cb_quotation_data["賣出一"], cb_stock_quotation_data["買進一"])
			if use_percentage:
				conversion_premium_rate *= 100.0
			if low_conversion_premium_rate_threshold is not None and conversion_premium_rate > low_conversion_premium_rate_threshold:
				continue
			if breakeven_threshold is not None and cb_quotation_data["成交"] > breakeven_threshold:
				continue

			issuing_date_days = self.__get_days(start_date_str=cb_publish_data["發行日期"])
			if issuing_date_days <= issuing_date_threshold:
				issuing_date_cb_dict[cb_id] = {
					"商品": cb_quotation_data["商品"],
					"日期": cb_publish_data["發行日期"],
					"天數": issuing_date_days,
					"溢價率": conversion_premium_rate,
					"成交": cb_quotation_data["成交"],
					"總量": cb_quotation_data["總量"],
					"發行張數": cb_publish_data["發行總面額"] / 100000,
				}
			convertible_date_days = self.__get_days(cb_summary_data["可轉換日"])
			if abs(convertible_date_days) <= convertible_date_threshold:
				convertible_date_cb_dict[cb_id] = {
					"商品": cb_quotation_data["商品"],
					"日期": cb_summary_data["可轉換日"],
					"天數": convertible_date_days,
					"溢價率": conversion_premium_rate,
					"成交": cb_quotation_data["成交"],
					"總量": cb_quotation_data["總量"],
					"發行張數": cb_publish_data["發行總面額"] / 100000,
				}
			maturity_date_days = self.__get_days(cb_publish_data["到期日期"])
			if maturity_date_days <= maturity_date_threshold:
				maturity_date_cb_dict[cb_id] = {
					"商品": cb_quotation_data["商品"],
					"日期": cb_publish_data["到期日期"],
					"天數": maturity_date_days,
					"溢價率": conversion_premium_rate,
					"成交": cb_quotation_data["成交"],
					"總量": cb_quotation_data["總量"],
					"發行張數": cb_publish_data["發行總面額"] / 100000,
				}
		return issuing_date_cb_dict, convertible_date_cb_dict, maturity_date_cb_dict


	def scrape_publish_detail(self, cb_id):
		# import pdb; pdb.set_trace()
		url = self.cb_publish[cb_id]["發行資料"]
		resp = self.__get_requests_module().get(url)
		if re.search("查無債券基本資料", resp.text) is not None:
			return None
		# print(resp.text)
		beautifulsoup_class = self.__get_beautifulsoup_class()
		soup = beautifulsoup_class(resp.text, "html.parser")

		table = soup.find_all("table", {"class": "hasBorder"})
# Beautiful Soup: 'ResultSet' object has no attribute 'find_all'?
# https://stackoverflow.com/questions/24108507/beautiful-soup-resultset-object-has-no-attribute-find-all
		table_trs = table[0].find_all("tr")
		cb_publish_detail_dict = {}
		# import pdb; pdb.set_trace()
		for tr in table_trs:
			tds = tr.find_all("td")
			for td in tds:
				# print(td.text)
				td_elem_list = list(map(lambda x: x.strip(), td.text.split("：", 1)))
				if len(td_elem_list) < 2:
					# print("ERROR: %s" % td.text)
					# import pdb; pdb.set_trace()
					continue
				cb_publish_detail_dict[td_elem_list[0]] = td_elem_list[1]
		# print("===============================================================================")
		# print(cb_publish_detail_dict)
		# print("===============================================================================")
		# print("本月受理轉(交)換之公司債張數: %s" % (cb_publish_detail_dict["本月受理轉(交)換之公司債張數"]))
		# print("最新轉(交)換價格: %s" % (cb_publish_detail_dict["最新轉(交)換價格"]))
		# print("最近轉(交)換價格生效日期: %s" % (cb_publish_detail_dict["最近轉(交)換價格生效日期"]))
		return cb_publish_detail_dict


	def __stock_info_profitability_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		MAX_ENTRY_COUNT = 4 * 5  # 4 quaters per year * 5 years
		data_dict = OrderedDict()
		table = driver.find_element("xpath", '//*[@id="oMainTable"]')
		trs = table.find_elements("tag name", "tr")

		table_row_start_index = 2
		title_list = []
		tds = trs[table_row_start_index].find_elements("tag name", "td")
		for td in tds[1:]:
			title_list.append(td.text)
		# import pdb; pdb.set_trace()
		for index, tr in enumerate(trs[table_row_start_index + 1:]):
			if index == MAX_ENTRY_COUNT: break
			tds = tr.find_elements("tag name", "td")
			td_text_list = []
			for td in tds[1:]:
				td_text_list.append(td.text)
			# import pdb; pdb.set_trace()
			data_dict[tds[0].text] = dict(zip(title_list, td_text_list))
		# import pdb; pdb.set_trace()
		# print("Data Count: %d" % len(data_dict.items()))
		return data_dict


	def __stock_info_revenue_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		MAX_ENTRY_COUNT = 12 * 3  # 12 months per year * 3 years
		data_dict = OrderedDict()
		table = driver.find_element("xpath", '//*[@id="oMainTable"]')
		trs = table.find_elements("tag name", "tr")

		table_row_start_index = 5
		title_list = []
		tds = trs[table_row_start_index].find_elements("tag name", "td")
		for td in tds[1:]:
			title_list.append(td.text)
		# import pdb; pdb.set_trace()
		for index, tr in enumerate(trs[table_row_start_index + 1:]):
			if index == MAX_ENTRY_COUNT: break
			tds = tr.find_elements("tag name", "td")
			td_text_list = []
			for td in tds[1:]:
				td_text_list.append(td.text)
			# import pdb; pdb.set_trace()
			data_dict[tds[0].text] = dict(zip(title_list, td_text_list))
		# import pdb; pdb.set_trace()
		return data_dict


	def __stock_info_earning_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		MAX_ENTRY_COUNT = 4 * 5  # 4 quaters per year * 5 years
		selenium_select_module = self.__get_selenium_select_module()
		list_box = selenium_select_module.Select(driver.find_element("xpath", '//*[@id="SysJustIFRAMEDIV"]/table[1]/tbody/tr/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/form/table/tbody/tr/td/table/tbody/tr[1]/td/select'))
		list_box.select_by_index(1)  # Replace '2' with the desired option's index
		time.sleep(3)

		# import pdb; pdb.set_trace()
		data_dict = OrderedDict()
		table = driver.find_element("xpath", '//*[@id="SysJustIFRAMEDIV"]/table[1]/tbody/tr/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/form/table/tbody/tr/td/table')
		trs = table.find_elements("tag name", "tr")

		table_row_start_index = 3
		title_list = []
		tds = trs[table_row_start_index].find_elements("tag name", "td")
		for td in tds[1:]:
			title_list.append(td.text)
		# import pdb; pdb.set_trace()
		for index, tr in enumerate(trs[table_row_start_index + 1:]):
			if index == MAX_ENTRY_COUNT: break
			tds = tr.find_elements("tag name", "td")
			td_text_list = []
			for td in tds[1:]:
				td_text_list.append(td.text)
			# import pdb; pdb.set_trace()
			data_dict[tds[0].text] = dict(zip(title_list, td_text_list))
		# import pdb; pdb.set_trace()
		return data_dict


	def __stock_info_cooperate_shareholding_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		data_dict = {}
		table = driver.find_element("xpath", '//*[@id="SysJustIFRAMEDIV"]/table[1]/tbody/tr/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/form/table/tbody/tr/td/table')
		trs = table.find_elements("tag name", "tr")

		table_row_start_index = 5
		title_tmp_list1 = []
		td1s = trs[table_row_start_index].find_elements("tag name", "td")
		for td in td1s:
			title_tmp_list1.append(td.text)
		title_tmp_list2 = []
		td2s = trs[table_row_start_index + 1].find_elements("tag name", "td")
		for td in td2s:
			title_tmp_list2.append(td.text)
		# import pdb; pdb.set_trace()
		title_list = []
		title_list.extend(list(map(lambda x: "%s%s" % (x, title_tmp_list1[1]), title_tmp_list2[1:5])))
		title_list.extend(list(map(lambda x: "%s%s" % (x, title_tmp_list1[2]), title_tmp_list2[5:9])))
		title_list.extend(list(map(lambda x: "%s%s" % (x, title_tmp_list1[3]), title_tmp_list2[9:])))
		# import pdb; pdb.set_trace()
		for tr in trs[table_row_start_index + 2:]:
			tds = tr.find_elements("tag name", "td")
			td_text_list = []
			for td in tds[1:]:
				td_text_list.append(td.text)
			# import pdb; pdb.set_trace()
			data_dict[tds[0].text] = dict(zip(title_list, td_text_list))
		# import pdb; pdb.set_trace()
		return data_dict


	# def __stock_info_margin_trading_scrapy_funcptr(self, driver):
	# 	data_dict = {}
	# 	table = driver.find_element("xpath", '//*[@id="SysJustIFRAMEDIV"]/table[1]/tbody/tr/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/table[1]/tbody/tr/td/table[6]')
	# 	trs = table.find_elements("tag name", "tr")
	# 	import pdb; pdb.set_trace()
	# 	mobj = re.search("融資融券", trs[0].find_element("tag name", "td").text)
	# 	if mobj is not None:
	# 		title_tmp_list1 = []
	# 		td1s = trs[1].find_elements("tag name", "td")
	# 		for td in td1s[1:3]:
	# 			title_tmp_list1.append(td.text)
	# 		title_tmp_list2 = []
	# 		td2s = trs[2].find_elements("tag name", "td")
	# 		for td in td2s[1:]:
	# 			title_tmp_list2.append(td.text)
	# 		# import pdb; pdb.set_trace()
	# 		title_list = []
	# 		title_list.extend(list(map(lambda x: "%s%s" % (title_tmp_list1[0], x), title_tmp_list2[1:7])))
	# 		title_list.extend(list(map(lambda x: "%s%s" % (title_tmp_list1[1], x), title_tmp_list2[7:])))
	# 		title_list.append(title_tmp_list2[-1])
	# 		# import pdb; pdb.set_trace()
	# 		for tr in trs[3:]:
	# 			tds = tr.find_elements("tag name", "td")
	# 			td_text_list = []
	# 			for td in tds[1:]:
	# 				td_text_list.append(td.text)
	# 			# import pdb; pdb.set_trace()
	# 			data_dict[tds[0].text] = dict(zip(title_list, td_text_list))
	# 	return data_dict


	def __stock_info_margin_trading_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		data_dict = {}
		table = driver.find_element("xpath", '//*[@id="SysJustIFRAMEDIV"]/table[1]/tbody/tr/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/form/table/tbody/tr/td/table')
		trs = table.find_elements("tag name", "tr")

		table_row_start_index = 5
		title_tmp_list1 = []
		td1s = trs[table_row_start_index].find_elements("tag name", "td")
		for td in td1s:
			title_tmp_list1.append(td.text)
		title_tmp_list2 = []
		td2s = trs[table_row_start_index + 1].find_elements("tag name", "td")
		for td in td2s:
			title_tmp_list2.append(td.text)
		# import pdb; pdb.set_trace()
		title_list = []
		title_list.extend(list(map(lambda x: "%s%s" % (title_tmp_list1[1], x), title_tmp_list2[1:8])))
		title_list.extend(list(map(lambda x: "%s%s" % (title_tmp_list1[2], x), title_tmp_list2[8:-1])))
		title_list.append("%s%s" % (title_tmp_list1[-1], title_tmp_list2[-1]))
		# import pdb; pdb.set_trace()
		for tr in trs[table_row_start_index + 2:]:
			tds = tr.find_elements("tag name", "td")
			td_text_list = []
			for td in tds[1:]:
				td_text_list.append(td.text)
			# import pdb; pdb.set_trace()
			data_dict[tds[0].text] = dict(zip(title_list, td_text_list))
		# import pdb; pdb.set_trace()
		return data_dict


	def __stock_info_major_inflow_outflow_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		data_dict = {}
		table = driver.find_element("xpath", '//*[@id="oMainTable"]')
		trs = table.find_elements("tag name", "tr")

		table_row_start_index = 6
		main_title_list = []
		tds = trs[table_row_start_index].find_elements("tag name", "td")
		for td in tds:
			main_title_list.append(td.text)
			data_dict[td.text] = {}
		table_row_start_index += 1
		title_list = []
		tds = trs[table_row_start_index].find_elements("tag name", "td")
		for td in tds:
			title_list.append(td.text)
		# import pdb; pdb.set_trace()
		table_row_start_index += 1
		table_row_end_index = table_row_start_index + 15
		for tr in trs[table_row_start_index:table_row_end_index]:
			tds = tr.find_elements("tag name", "td")
			td_text_list = []
			for td in tds:
				td_text_list.append(td.text)
			# import pdb; pdb.set_trace()
			data_dict[main_title_list[0]][tds[0].text] = dict(zip(title_list[1:5], td_text_list[1:5]))
			data_dict[main_title_list[1]][tds[5].text] = dict(zip(title_list[6:], td_text_list[6:]))
		table_row_start_index = table_row_end_index
		# import pdb; pdb.set_trace()
		for index in range(table_row_start_index, table_row_start_index + 2):
			tds = trs[index].find_elements("tag name", "td")
			data_dict[main_title_list[0]][tds[0].text] = tds[1].text
			data_dict[main_title_list[1]][tds[2].text] = tds[3].text
		# import pdb; pdb.set_trace()
		return data_dict


	def __stock_info_balance_sheet_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		data_dict = {}
		table = driver.find_element("xpath", '//*[@id="oMainTable"]')
		divs = table.find_elements("tag name", "div")

		table_row_start_index = 4
		period_list = []
		spans = divs[table_row_start_index].find_elements("tag name", "span")
		for span in spans[1:]:
			data_dict[span.text] = {}
			period_list.append(span.text)
		table_row_start_index += 2
		for div in divs[table_row_start_index:]:
			spans = div.find_elements("tag name", "span")
			title = spans[0].text
			for index, span in enumerate(spans[1:]):
				data_dict[period_list[index]][title] = span.text
		# import pdb; pdb.set_trace()
		return data_dict


	def __stock_info_cash_flow_statement_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		data_dict = {}
		table = driver.find_element("xpath", '//*[@id="oMainTable"]')
		divs = table.find_elements("tag name", "div")

		table_row_start_index = 4
		period_list = []
		spans = divs[table_row_start_index].find_elements("tag name", "span")
		for span in spans[1:]:
			data_dict[span.text] = {}
			period_list.append(span.text)
		table_row_start_index += 2
		for div in divs[table_row_start_index:]:
			spans = div.find_elements("tag name", "span")
			title = spans[0].text
			for index, span in enumerate(spans[1:]):
				data_dict[period_list[index]][title] = span.text
		# import pdb; pdb.set_trace()
		return data_dict


	def __stock_info_financial_ratio_statement_scrapy_funcptr(self, driver):
		# import pdb; pdb.set_trace()
		data_dict = {}
		table = driver.find_element("xpath", '//*[@id="oMainTable"]')
		divs = table.find_elements("tag name", "div")

		table_row_start_index = -1
		period_list = None
		# import pdb; pdb.set_trace()
		for index, div in enumerate(divs):
			# print(div.text)
			spans = div.find_elements("tag name", "span")
			if len(spans) != 0 and re.match("期別", spans[0].text):
				period_list = []
				for span in spans[1:]:
					data_dict[span.text] = {}
					period_list.append(span.text)
				table_row_start_index = index + 1
				break
		if period_list is None:
			raise RuntimeError('Fails to find the "期別" field in the table')
		table_column_len = len(period_list) + 1
		for div in divs[table_row_start_index:]:
			spans = div.find_elements("tag name", "span")
			if len(spans) != table_column_len:
				continue
			elif (re.match("期別", spans[0].text) is not None) or (re.match("種類", spans[0].text) is not None):
				continue
			title = spans[0].text
			for index, span in enumerate(spans[1:]):
				data_dict[period_list[index]][title] = span.text
		# import pdb; pdb.set_trace()
		return data_dict

	def __calculate_stock_info_daily_update_time(self, data_scrapy_date):
		daily_new_data_date = data_scrapy_date + timedelta(days=1)
		return daily_new_data_date


	def __calculate_stock_info_monthly_update_time(self, data_scrapy_date):
		monthly_new_data_date = None
		if data_scrapy_date.day <= 10:
			monthly_new_data_date = datetime(data_scrapy_date.year, data_scrapy_date.month, 10)
		else:
			year = data_scrapy_date.year
			month = data_scrapy_date.month + 1
			if month > 12:
				month -= 12
				year += 1
			monthly_new_data_date = datetime(year, month, 10)
		return monthly_new_data_date


	def __calculate_stock_info_quarterly_update_time(self, data_scrapy_date):
		statement_release_date_check_list = []
		for date in self.STATEMENT_RELEASE_DATE_LIST:
			statement_release_date_check_list.append(datetime(data_scrapy_date.year, date[0], date[1]))
		statement_release_date_check_list.append(datetime(data_scrapy_date.year + 1, self.STATEMENT_RELEASE_DATE_LIST[0][0], self.STATEMENT_RELEASE_DATE_LIST[0][1]))
		check_index = len(list(filter(lambda x: x < data_scrapy_date, statement_release_date_check_list)))
		quarterly_new_data_date = statement_release_date_check_list[check_index]
		return quarterly_new_data_date


	def __calculate_stock_info_yearly_update_time(self, data_scrapy_date):
		statement_release_date_check_list = []
		statement_release_date_check_list.append(datetime(data_scrapy_date.year, self.STATEMENT_RELEASE_DATE_LIST[0][0], self.STATEMENT_RELEASE_DATE_LIST[0][1]))
		statement_release_date_check_list.append(datetime(data_scrapy_date.year + 1, self.STATEMENT_RELEASE_DATE_LIST[0][0], self.STATEMENT_RELEASE_DATE_LIST[0][1]))
		check_index = len(list(filter(lambda x: x < data_scrapy_date, statement_release_date_check_list)))
		yearly_new_data_date = statement_release_date_check_list[check_index]
		return yearly_new_data_date


	def calculate_stock_info_update_time(self, data_scrapy_timestr, update_frequency):
		# import pdb; pdb.set_trace()
		data_scrapy_date = datetime.strptime(data_scrapy_timestr, "%Y/%m/%d")
		return (self.STOCK_INFO_UPDATE_TIME_FUNCPTR_DICT[update_frequency])(data_scrapy_date)
		# print("Now: %s\nDaily: %s\nMonthly: %s\nQuarterly: %s\nYearly: %s" % (data_scrapy_timestr, str(daily_new_data_date), str(monthly_new_data_date), str(quarterly_new_data_date), str(yearly_new_data_date)))


	def scrape_stock_info(self, cb_id, from_file=True):
		data_dict = None
		# import pdb; pdb.set_trace()
		driver = self.__get_web_driver()
		cb_stock_id = cb_id[:4]
		# import pdb; pdb.set_trace()
		data_filepath = os.path.join(self.xcfg["cb_data_folderpath"], "%s.txt" % cb_stock_id)
		if from_file:
			if self.__check_file_exist(data_filepath):
				with open(data_filepath, "r", encoding='utf-8') as f:
					data_dict = json.load(f)
		if data_dict is None:
			data_dict = {
				"time": {}, 
				"content": {},
			}
		try:
			cur_datetime = datetime.now()
			total_start_time = time.time()
			data_time_dict = data_dict["time"]
			data_content_dict = data_dict["content"]
			for scrapy_key, scrapy_funcptr in self.STOCK_INFO_SCRAPY_FUNCPTR_DICT.items():
# Check it's required to scrape the data
				need_scrapy = True
				if scrapy_key in data_time_dict:
					data_time_str = data_time_dict[scrapy_key]
					update_frequency = self.STOCK_INFO_SCRAPY_UPDATE_FREQUENCY_DICT[scrapy_key]
					new_update_time = self.calculate_stock_info_update_time(data_time_str, update_frequency)
					if new_update_time > cur_datetime:
						need_scrapy = False
				if need_scrapy:
					url = self.STOCK_INFO_SCRAPY_URL_FORMAT_DICT[scrapy_key] % cb_stock_id
					print("Scrape %s......" % scrapy_key)
					start_time = time.time()
					driver.get(url)
					time.sleep(5)
					data_content_dict[scrapy_key] = scrapy_funcptr(driver)
					end_time = time.time()
					print("Scrape %s...... Done in %d seconds" % (scrapy_key, (end_time - start_time)))
					data_time_dict[scrapy_key] = cur_datetime.strftime("%Y/%m/%d %H:%M:%S")
			print(data_dict)
			total_end_time = time.time()
			print("Scrape All...... Done in %d seconds" % (total_end_time - total_start_time))
# Writing to file
			# import pdb; pdb.set_trace()
			if from_file:
				with open(data_filepath, "w", encoding='utf-8') as f:
					json.dump(data_dict, f, indent=3, ensure_ascii=False)
		except Exception as e:
			print(e)
# Close the web driver in the __exit__ function
		# finally:
		# 	driver.close()
		return data_dict


	def calculate_cb_monthly_convert_data_table_month(self):
		today = datetime.today()
		year = today.year - 1911
		month = today.month
		day = today.day
		if day >= 11:
			month -= 1
		else:
			month -= 2
		if month <= 0:
			month += 12
			year -= 1
		# filename = "%s%d%02d" % (self.self.DEFAULT_CB_MONTHLY_CONVERT_DATA_FILENAME_PREFIX, year, month)
		# # filepath = os.path.join(self.xcfg["cb_data_folderpath"], filename)
		# # return (not self.__check_file_exist(filepath))
		table_month = "%d%02d" % (year, month)
		return table_month


	def scrape_cb_monthly_convert_data(self):
		data_dict = {}
		driver = self.__get_web_driver()
		url = "https://www.tdcc.com.tw/portal/zh/QStatWAR/indm004"
		try:		
			driver.get(url)
			time.sleep(5)
			btn = driver.find_element("xpath", '//*[@id="form1"]/table/tbody/tr[4]/td/input')
			btn.click()
			time.sleep(5)
			table = driver.find_element("xpath", '//*[@id="body"]/div/main/div[6]/div/table')
# Check the table time
			# import pdb; pdb.set_trace()
			span = driver.find_element("xpath", '//*[@id="body"]/div/main/div[5]/span')
			mobj = re.search(".+([\d]{5})", span.text)
			if mobj is None:
				raise RuntimeError("Fail to find the month of the table")
			table_month = mobj.group(1)
			filename = self.DEFAULT_CB_MONTHLY_CONVERT_DATA_FILENAME_PREFIX + table_month
			filepath = os.path.join(self.xcfg["cb_data_folderpath"], filename)
			# import pdb; pdb.set_trace()
			if self.__check_file_exist(filepath):
				# print ("The file[%s] already exist !!!" % filepath)
				with open(filepath, "r", encoding='utf8') as f:
					data_dict = json.load(f)
			else:
# thead
				print ("The file[%s] does NOT exist. Scrape the data from website" % filepath)
				table_head = table.find_element("tag name", "thead")
				trs = table_head.find_elements("tag name", "tr")
				table_title_list = []
				ths = trs[1].find_elements("tag name", "th")
				for th in ths:
					table_title_list.append(th.text)
				ths = trs[0].find_elements("tag name", "th")
				for th in ths[1:]:
					table_title_list.append(th.text.split('\n')[0])
				# print(table_title_list)
# tbody
				table_body = table.find_element("tag name", "tbody")
				trs = table_body.find_elements("tag name", "tr")
				for tr in trs:
					tds = tr.find_elements("tag name", "td")
					td_text_list = []
					for td in tds:
						td_text_list.append(td.text.replace(",",""))
					# print(", ".join(td_text_list))
					data_dict[td_text_list[0]] = dict(zip(table_title_list[1:], td_text_list[1:]))
				# time.sleep(5)
				# import pdb; pdb.set_trace()
# Writing to file
				with open(filepath, "w", encoding='utf-8') as f:
					json.dump(data_dict, f, indent=3, ensure_ascii=False)
		except Exception as e:
			print ("Exception occurs while scraping [%s], due to: %s" % (url, str(e)))
			raise e
		# finally:
		# 	driver.close()
		# for key, value in data_dict.items():
		# 	value_str_list = list(map(lambda x: "%s(%s)" % (x[0], x[1]), value.items()))
		# 	print("%s: %s" % (key, ", ".join(value_str_list)))
		return data_dict


	def get_cb_monthly_convert_data(self, table_month=None):
		# import pdb; pdb.set_trace()
		filepath = None
		scrapy_data_dict = None
		if table_month is None:
			table_month = self.calculate_cb_monthly_convert_data_table_month()
		filename = self.DEFAULT_CB_MONTHLY_CONVERT_DATA_FILENAME_PREFIX + table_month
		filepath = os.path.join(self.xcfg["cb_data_folderpath"], filename)
		if self.__check_file_exist(filepath):
			with open(filepath, 'r', encoding='utf-8') as f:
				scrapy_data_dict = json.load(f)
		if scrapy_data_dict is None:
			scrapy_data_dict = self.scrape_cb_monthly_convert_data()
			if table_month is not None:
				if not self.__check_file_exist(filepath):
					raise ValueError("The data of %s is NOT found" % os.path.basename(filepath))
# Fails to read from the TXT file
		# url = "https://m.tdcc.com.tw/tcdata/sm/bimon92.txt"
		# data_filename = "monthly_convert_data.txt"
		# resp = self.__get_wget_module().download(url, data_filename)
		# data_filepath = os.path.join(os.getcwd(), data_filename)
		# # import pdb; pdb.set_trace()
		# if not self.__check_file_exist(data_filepath):
		# 	raise RuntimeError("Fails to download the file: %s" % data_filepath)
		# # import pdb; pdb.set_trace()
		# with open(data_filepath, "rb") as f:
		# 	for line in f:
		# 		print(line)
		return scrapy_data_dict
			

	def search_cb_mass_convert(self, table_month=None, mass_convert_threshold=-10.0):
		# import pdb; pdb.set_trace()
		convert_cb_dict = self.get_cb_monthly_convert_data(table_month)
		if not self.xcfg['cb_all']:
			convert_cb_dict = dict(filter(lambda x: x[0] in self.cb_id_list, convert_cb_dict.items()))
		mass_convert_cb_dict = dict(filter(lambda x: float(x[1]["增減百分比"]) < mass_convert_threshold, convert_cb_dict.items()))
		return mass_convert_cb_dict


	def get_publish_detail(self, cb_id):
		if cb_id not in self.cb_publish_detail:
			self.cb_publish_detail[cb_id] = self.scrape_publish_detail(cb_id)
		return self.cb_publish_detail[cb_id]


	def search_multiple_publish(self):
		multiple_publish_dict = {}
		for cb_stock_id in self.cb_stock_id_list:
			regex = re.compile(cb_stock_id)
			filter_cb_id_list = list(filter(regex.match, self.cb_id_list))
			if len(filter_cb_id_list) > 1:
				multiple_publish_dict[cb_stock_id] = filter_cb_id_list
		return multiple_publish_dict


	@property
	def CBSummary(self):
		assert self.cb_summary is not None, "cb_summary should NOT be NONE"
		return self.cb_summary


	@property
	def CBPublish(self):
		assert self.cb_publish is not None, "cb_publish should NOT be NONE"
		return self.cb_publish


	@property
	def CanScrape(self):
		return self.can_scrape


	def search(self):
		# data_dict_summary = self.__read_cb_summary()
		# # print (data_dict_summary)
		quotation_data_dict = self.__read_cb_quotation()
		stock_quotation_data_dict = self.__read_cb_stock_quotation()
		if self.xcfg['cb_all']:
			self.check_data_source(quotation_data_dict, stock_quotation_data_dict)
		print("\n*****************************************************************\n")

		# print (quotation_data_dict)
		# print(self.calculate_internal_rate_of_return(quotation_data_dict))
		irr_dict = self.get_positive_internal_rate_of_return(quotation_data_dict)
		if bool(irr_dict):
			print("=== 年化報酬率 ==================================================")
			title_list = ["年化報酬率", "賣出一", "到期日",]
			print("  ===> %s" % ", ".join(title_list))
			for irr_key, irr_data in irr_dict.items():
				print ("%s[%s]: %.2f  %.2f  %s" % (irr_data["商品"], irr_key, float(irr_data["年化報酬率"]), float(irr_data["賣出一"]), irr_data["到期日"]))
			print("=================================================================\n")
		premium_dict = self.get_negative_premium(quotation_data_dict, stock_quotation_data_dict)
		if bool(premium_dict):
			print("=== 溢價率(套利) ================================================")
			title_list = ["溢價率", "融資餘額", "融券餘額",]
			print("  ===> %s" % ", ".join(title_list))
			# import pdb; pdb.set_trace()
			for premium_key, premium_data in premium_dict.items():
				print ("%s[%s]: %.2f  %d  %d" % (premium_data["商品"], premium_key, float(premium_data["溢價率"]), premium_data["融資餘額"], premium_data["融券餘額"]))
			print("=================================================================\n")
		stock_premium_dict = self.get_absolute_stock_premium(quotation_data_dict, stock_quotation_data_dict)
		if bool(stock_premium_dict):
			print("=== 股票溢價率 ==================================================")
			title_list = ["股票溢價率",]
			print("  ===> %s" % ", ".join(title_list))
			for stock_premium_key, stock_premium_data in stock_premium_dict.items():
				print ("%s[%s]: %.2f" % (stock_premium_data["商品"], stock_premium_key, float(stock_premium_data["股票溢價率"])))
			print("=================================================================\n")
		cb_dict = self.get_low_premium_and_breakeven(quotation_data_dict, stock_quotation_data_dict)
		if bool(cb_dict):
			print("=== 低溢價且保本 ================================================")
			title_list = ["溢價率", "成交", "賣出一", "到期日",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in cb_dict.items():
				print ("%s[%s]: %.2f  %.2f  %.2f  %s" % (cb_data["商品"], cb_key, float(cb_data["溢價率"]), float(cb_data["成交"]), float(cb_data["賣出一"]), cb_data["到期日"]))
			print("=================================================================\n")
		issuing_date_cb_dict, convertible_date_cb_dict, maturity_date_cb_dict = self.search_cb_opportunity_dates(quotation_data_dict, stock_quotation_data_dict)
		if bool(issuing_date_cb_dict):
			print("=== 近發行日期 ==================================================")
			title_list = ["日期", "天數", "溢價率", "成交", "總量", "發行張數",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in issuing_date_cb_dict.items():
				print ("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d" % (cb_data["商品"], cb_key, cb_data["日期"], int(cb_data["天數"]), float(cb_data["溢價率"]), float(cb_data["成交"]), int(cb_data["總量"]), int(cb_data["發行張數"])))
				cb_publish_detail_dict = self.get_publish_detail(cb_key)
				print(" *************")
				if cb_publish_detail_dict is None:
					print("  查無債券基本資料......")
				else:
					print("  本月受理轉(交)換之公司債張數: %s" % (cb_publish_detail_dict["本月受理轉(交)換之公司債張數"]))
					print("  最新轉(交)換價格: %s" % (cb_publish_detail_dict["最新轉(交)換價格"]))
					# print("  最近轉(交)換價格生效日期: %s" % (cb_publish_detail_dict["最近轉(交)換價格生效日期"]))
				print(" *************")
			print("=================================================================\n")
		if bool(convertible_date_cb_dict):
			print("=== 近可轉換日 ==================================================")
			title_list = ["日期", "天數", "溢價率", "成交", "總量", "發行張數",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in convertible_date_cb_dict.items():
				print ("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d" % (cb_data["商品"], cb_key, cb_data["日期"], int(cb_data["天數"]), float(cb_data["溢價率"]), float(cb_data["成交"]), int(cb_data["總量"]), int(cb_data["發行張數"])))
				cb_publish_detail_dict = self.get_publish_detail(cb_key)
				print(" *************")
				print("  本月受理轉(交)換之公司債張數: %s" % (cb_publish_detail_dict["本月受理轉(交)換之公司債張數"]))
				print("  最新轉(交)換價格: %s" % (cb_publish_detail_dict["最新轉(交)換價格"]))
				print("  最近轉(交)換價格生效日期: %s" % (cb_publish_detail_dict["最近轉(交)換價格生效日期"]))
				print(" *************")
			print("=================================================================\n")
		if bool(maturity_date_cb_dict):
			print("=== 近到期日期 ==================================================")
			title_list = ["日期", "天數", "溢價率", "成交", "總量", "發行張數",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in maturity_date_cb_dict.items():
				print ("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d" % (cb_data["商品"], cb_key, cb_data["日期"], int(cb_data["天數"]), float(cb_data["溢價率"]), float(cb_data["成交"]), int(cb_data["總量"]), int(cb_data["發行張數"])))
				cb_publish_detail_dict = self.get_publish_detail(cb_key)
				print(" *************")
				print("  本月受理轉(交)換之公司債張數: %s" % (cb_publish_detail_dict["本月受理轉(交)換之公司債張數"]))
				print("  最新轉(交)換價格: %s" % (cb_publish_detail_dict["最新轉(交)換價格"]))
				# print("  最近轉(交)換價格生效日期: %s" % (cb_publish_detail_dict["最近轉(交)換價格生效日期"]))
				print(" *************")
			print("=================================================================\n")
		mass_convert_cb_dict = self.search_cb_mass_convert()
		if bool(mass_convert_cb_dict):
			print("=== CB大量轉換 ==================================================")
			title_list = ["增減百分比", "前月底保管張數", "本月底保管張數", "發行張數",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in mass_convert_cb_dict.items():
				print ("%s[%s]:  %.2f  %d  %d  %d" % (cb_data["名稱"], cb_key, float(cb_data["增減百分比"]), int(cb_data["前月底保管張數"]), int(cb_data["本月底保管張數"]), int(cb_data["發行張數"])))
		# multiple_publish_dict = self.search_multiple_publish()
		# if bool(multiple_publish_dict):
		# 	print("=== 多次發行 ==================================================")
		# 	for key, data in multiple_publish_dict.items():
		# 		print("***** %s *****" % key)
		# 		cb_id_list = data
		# 		for cb_id in cb_id_list:
		# 			# import pdb; pdb.set_trace()
		# 			cb_publish_data = self.CBPublish[cb_id]
		# 			print("%s[%s]: %s  %s  %d  %d" % (cb_publish_data["債券簡稱"], cb_id, cb_publish_data["發行日期"], cb_publish_data["到期日期"], cb_publish_data["年期"], cb_publish_data["發行總面額"] / 100000))
				# print("\n")


	def __get_cb_list_from_file(self):
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(self.xcfg['cb_list_filepath']):
			raise RuntimeError("The file[%s] does NOT exist" % self.xcfg['cb_list_filepath'])
		self.xcfg["cb_list"] = []
		with open(self.xcfg['cb_list_filepath'], 'r') as fp:
			for line in fp:
				if line.startswith("#"): continue
				line = line.strip("\n")
				if len(line) == 0: continue
				self.xcfg["cb_list"].append(line)


	def display(self):
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '到期日']
		quotation_data_dict = self.__read_cb_quotation()
		stock_quotation_data_dict = self.__read_cb_stock_quotation()
		if self.xcfg['cb_all']:
			self.check_data_source(quotation_data_dict, stock_quotation_data_dict)

		# import pdb; pdb.set_trace()
		irr_dict = self.calculate_internal_rate_of_return(quotation_data_dict, use_percentage=True)
		premium_dict = self.calculate_premium(quotation_data_dict, stock_quotation_data_dict, use_percentage=True)
		stock_premium_dict = self.calculate_stock_premium(quotation_data_dict, stock_quotation_data_dict, use_percentage=True)
		for cb_id in self.cb_id_list:
			cb_stock_id = cb_id[:4]
			summary_data = self.cb_summary[cb_id]
			publish_data = self.cb_publish[cb_id]
			quotation_data = quotation_data_dict[cb_id]
			stock_quotation_data = stock_quotation_data_dict[cb_stock_id]
			irr_data = irr_dict[cb_id]
			premium_data = premium_dict[cb_id]
			stock_premium_data = stock_premium_dict[cb_id]
			scrapy_data = self.scrape_stock_info(cb_stock_id)
			print("%s[%s]:" % (quotation_data["商品"], cb_id))
			print(" %s" % "  ".join(["溢價率", "成交", "賣出一",]))
			try:
				print(" %.2f  %.2f  %.2f" % (float(premium_data["溢價率"]), float(quotation_data["成交"]), float(quotation_data["賣出一"])))
			except TypeError:
				print(" %s" % ("  ".join([str(premium_data["溢價率"]), str(quotation_data["成交"]), str(quotation_data["賣出一"])])))
			print(" %s" % "  ".join(["發行日期", "年期", "發行總張數",]))
			try:
				print(" %s  %s  %d" % (publish_data["發行日期"], publish_data["年期"], int(publish_data["發行總面額"]) /100000))
			except TypeError:
				print(" %s" % ("  ".join([str(premium_data["溢價率"]), str(quotation_data["成交"]), str(quotation_data["賣出一"])])))


	def print_filepath(self):
		print("************** File Path **************")
		for key, value in self.filepath_dict.items():
			print("%s: %s" % (key, value))


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
	parser.add_argument('-a', '--all', required=False, action='store_true', help='Check all CBs.')
	parser.add_argument('-s', '--search', required=False, action='store_true', help='Select targets based on the search rule.')
	parser.add_argument('-d', '--display', required=False, action='store_true', help='Display specific targets.')
	parser.add_argument('--cb_list', required=False, help='The list of specific CB targets.')
	parser.add_argument('--print_filepath', required=False, action='store_true', help='Print the filepaths used in the process and exit.')
	args = parser.parse_args()

	cfg = {
		# "cb_list": "62822,62791,61965,54342,33881"
	}
	if args.all:
		cfg['cb_all'] = args.all
	if args.cb_list:
		cfg['cb_list'] = args.cb_list
		if 'cb_all' in cfg:
			print("The 'all' flag is ignored...")
			cfg['cb_all'] = False
	with ConvertibleBondAnalysis(cfg) as obj:
		# obj.calculate_stock_info_update_time("2024/03/25")
		# data_dict = obj.scrape_stock_info("2330")
		if args.print_filepath:
			obj.print_filepath()
			sys.exit(0)
		print(data_dict)
		if args.search:
			obj.search()
		# import pdb; pdb.set_trace()
		if args.display:
			obj.display()


# 	from selenium import webdriver
# 	import time
# 	# from selenium.webdriver.common.by import By



# 	driver = webdriver.Chrome("C:\chromedriver.exe")
# 	driver.get("https://www.tdcc.com.tw/portal/zh/QStatWAR/indm004")
# 	time.sleep(5)
# 	btn = driver.find_element("xpath", '//*[@id="form1"]/table/tbody/tr[4]/td/input')
# 	btn.click()
# 	time.sleep(5)
# 	table = driver.find_element("xpath", '//*[@id="body"]/div/main/div[6]/div/table')
# 	data_dict = {}
# # thead
# 	table_head = table.find_element("tag name", "thead")
# 	trs = table_head.find_elements("tag name", "tr")
# 	table_title_list = []
# 	ths = trs[1].find_elements("tag name", "th")
# 	for th in ths:
# 		table_title_list.append(th.text)
# 	ths = trs[0].find_elements("tag name", "th")
# 	for th in ths[1:]:
# 		table_title_list.append(th.text.split('\n')[0])
# 	print(table_title_list)
# # tbody
# 	table_body = table.find_element("tag name", "tbody")
# 	trs = table_body.find_elements("tag name", "tr")
# 	for tr in trs:
# 		tds = tr.find_elements("tag name", "td")
# 		td_text_list = []
# 		for td in tds:
# 			td_text_list.append(td.text)
# 		# print(", ".join(td_text_list))
# 		data_dict[td_text_list[0]] = dict(zip(table_title_list[1:], td_text_list[1:]))
# 	time.sleep(5)
# 	driver.close()

# 	for key, value in data_dict.items():
# 		value_str_list = list(map(lambda x: "%s(%s)" % (x[0], x[1]), value.items()))
# 		print("%s: %s" % (key, ", ".join(value_str_list)))

	# import pdb; pdb.set_trace()
	# import requests
	# image_url = "https://m.tdcc.com.tw/tcdata/sm/bimon92.txt"
	# # URL of the image to be downloaded is defined as image_url
	# r = requests.get(image_url) # create HTTP response object
	  
	# # send a HTTP request to the server and save
	# # the HTTP response in a response object called r
	# with open("python_logo.png",'wb') as f:
	  
	#     # Saving received content as a png file in
	#     # binary format
	  
	#     # write the contents of the response (r.content)
	#     # to a new file in binary mode.
	#     f.write(r.content)

	# driver.get("https://www.tpex.org.tw/web/bond/publish/convertible_bond_search/memo.php?l=zh-tw")
	# # driver.get("https://mops.twse.com.tw/mops/web/t120sg01?TYPEK=&bond_id=45552&bond_kind=5&bond_subn=%24M00000001&bond_yrn=2&come=2&encodeURIComponent=1&firstin=ture&issuer_stock_code=4555&monyr_reg=202302&pg=&step=0&tg=k_code=4555&monyr_reg=202302&pg=&step=0&tg=")
	# # driver.get("https://www.tdcc.com.tw/portal/zh/QStatWAR/indm004")
	# driver.get("https://concords.moneydj.com/Z/ZC/ZCX/ZCX_2330.djhtm")
	# time.sleep(5)
	# # #找到輸入框
	# # # element = driver.find_element_by_name("q");
	# # #form1 > table > tbody > tr:nth-child(4) > td > input[type=submit]
	# # link = driver.find_element("xpath", '//*[@id="table01"]/center/table[2]/tbody/tr[46]/td/a')
	# # # btn = driver.find_element("xpath", '/html/body/div[1]/div[1]/div/main/div[4]/form/table/tbody/tr[4]/td/input')
	# # time.sleep(5)
	# # btn.click()
	# table = driver.find_element("xpath", '//*[@id="SysJustIFRAMEDIV"]/table[1]/tbody/tr/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td/table[1]/tbody/tr/td/table[6]')
	# trs = table.find_elements("tag name", "tr")
	# # import pdb; pdb.set_trace()
	# data_dict = {}
	# mobj = re.search("融資融券", trs[0].find_element("tag name", "td").text)
	# if mobj is not None:
	# 	title_tmp_list1 = []
	# 	td1s = trs[1].find_elements("tag name", "td")
	# 	for td in td1s[1:3]:
	# 		title_tmp_list1.append(td.text)
	# 	title_tmp_list2 = []
	# 	td2s = trs[2].find_elements("tag name", "td")
	# 	for td in td2s[1:]:
	# 		title_tmp_list2.append(td.text)
	# 	# import pdb; pdb.set_trace()
	# 	title_list = []
	# 	title_list.extend(list(map(lambda x: "%s%s" % (title_tmp_list1[0], x), title_tmp_list2[1:7])))
	# 	title_list.extend(list(map(lambda x: "%s%s" % (title_tmp_list1[1], x), title_tmp_list2[7:])))
	# 	title_list.append(title_tmp_list2[-1])
	# 	# import pdb; pdb.set_trace()
	# 	for tr in trs[3:]:
	# 		tds = tr.find_elements("tag name", "td")
	# 		td_text_list = []
	# 		for td in tds[1:]:
	# 			td_text_list.append(td.text)
	# 		# import pdb; pdb.set_trace()
	# 		data_dict[tds[0].text] = dict(zip(title_list, td_text_list))
	# # import pdb; pdb.set_trace()
	# print(data_dict)
	# 		# print("%s\n" % (", ".join(td_text_list)))


	# # #輸入內容
	# # element.send_keys("hello world");
	# # #提交表單
	# # element.submit();
	# driver.close()

		# # # web_driver = self.__get_web_driver()
		# # # driver.get(url)
		# # resp = self.__get_requests_module().get(url)
		# # beautifulsoup_class = getattr(self.__get_beautifulsoup_module(), "beautifulsoup")
		# # soup = beautifulsoup_class.BeautifulSoup(resp.text)
		# resp = self.__get_requests_module().get(url)
		# # print(resp.text)
		# # bs4_module = __import__("bs4")
		# beautifulsoup_class = self.__get_beautifulsoup_class()
		# soup = beautifulsoup_class(resp.text, "html.parser")
