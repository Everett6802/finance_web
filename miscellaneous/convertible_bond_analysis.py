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
from dateutil.relativedelta import relativedelta
import math
import csv
import collections
import time
import json
import copy
from collections import OrderedDict


class RetryException(Exception): pass

def retry(max_retry=3, sleep_time=10):
	def decorator(func):
		def wrapper(*args, **kwargs):
			for index in range(max_retry):
				try:
					res = func(*args, **kwargs)
					return res
				except RetryException as e:
					print("Error occurs, due to: %s" % str(e))
					print("Sleep %d seconds, then retry...... %d" % (sleep_time, index + 1))
					time.sleep(sleep_time)
		return wrapper
	return decorator


class ConvertibleBondAnalysis(object):

	DEFAULT_CONFIG_FOLDERPATH =  "C:\\Users\\%s" % os.getlogin()
	DEFAULT_TRACKED_CB_LIST_FILENAME = "convertible_bond_list.txt"

	DEFAULT_CB_FOLDERPATH =  "C:\\可轉債"
	DEFAULT_CB_DATA_FOLDERNAME =  "Data"
	DEFAULT_CB_SUMMARY_FILENAME = "可轉債總表"	
	DEFAULT_CB_SUMMARY_FULL_FILENAME = "%s.csv" % DEFAULT_CB_SUMMARY_FILENAME
	PUBLISH_CB_LIST_URL = "https://www.tpex.org.tw/zh-tw/bond/issue/cbond/listed.html"
	

	@classmethod
	def __int(cls, value):
		return int(value) if (value is not None) else value


	@classmethod
	def __float(cls, value, digits=2):
		if value is None: return value
		if digits is not None:
			# import pdb; pdb.set_trace()
			str_format = "{:." + str(digits) + "f}"
			return str_format.format(value)
		return float(value)


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
				print("%s: %s" % (errno.errorcode[exception.errno], os.strerror(exception.errno)))
				raise
			check_exist = False
		return check_exist


	@classmethod
	def __strptime(cls, datetime_str, format_list=["%m/%d/%Y", "%Y/%m/%d",], year_transform=False, year_transform_splitter="/"):
		def transform_date_str(date_str):
			year_str, rest_str = date_str.split(year_transform_splitter, 1)
			new_year_str = str(int(year_str) + 1911)
			new_date_str = new_year_str + "/" + rest_str
			return new_date_str
		if year_transform:
			datetime_str = transform_date_str(datetime_str)

		for fmt in format_list:
			try:
				datetime_obj = datetime.strptime(datetime_str, fmt)
				return datetime_obj
			except ValueError:
				pass
		raise ValueError("Time string[%s] => Incorrect time format" % (datetime_str, str(format_list)))


	@classmethod
	def __get_days(cls, end_date_str=None, start_date_str=None):
		# import pdb; pdb.set_trace()
		time_format = "%m/%d/%Y"  # "%Y/%m/%d"
		end_date = None
		start_date = None
		if end_date_str is not None:
			# try:
			# 	end_date = datetime.strptime(end_date_str, "%Y/%m/%d")
			# except ValueError:
			# 	try:
			# 		end_date = datetime.strptime(end_date_str, "%m/%d/%Y")
			# 	except ValueError as e:
			# 		raise e
			end_date = cls.__strptime(end_date_str)
		else:
			end_date = datetime.fromordinal(date.today().toordinal())
		if start_date_str is not None:
			# try:
			# 	start_date = datetime.strptime(start_date_str, "%Y/%m/%d")
			# except ValueError:
			# 	try:
			# 		start_date = datetime.strptime(start_date_str, "%m/%d/%Y")
			# 	except ValueError as e:
			# 		raise e
			start_date = cls.__strptime(start_date_str)
		else:
			start_date = datetime.fromordinal(date.today().toordinal())
		# end_date = datetime.strptime(end_date_str, time_format) if end_date_str is not None else datetime.fromordinal(date.today().toordinal())
		# start_date = datetime.strptime(start_date_str, time_format) if start_date_str is not None else datetime.fromordinal(date.today().toordinal())
		days = int((end_date - start_date).days) + 1
		return days


	@classmethod
	def __get_conversion_ratio(cls, conversion_price):
		return 100.0 / float(conversion_price)


	@classmethod
	def __get_conversion_premium_rate(cls, conversion_price, bond_price, share_price):
		conversion_ratio = cls.__get_conversion_ratio(conversion_price)
		conversion_value = float(conversion_ratio) * float(share_price)
		conversion_premium_rate = (float(bond_price) - float(conversion_value)) / conversion_value
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


	@classmethod
	def __get_file_modification_date(cls, filepath):
		if not cls.__check_file_exist(filepath):
			raise ValueError("The file[%s] does NOT exist" % filepath)
		# print("filepath: %s" % filepath)
		# create_time = os.path.getctime(filepath)
		# print(create_time)
		modification_time = os.path.getmtime(filepath)
		# print(modification_time)
		modification_date = datetime.fromtimestamp(modification_time)
		return modification_date


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
			"cb_list": None,
			"cb_ignore_list": None,
			"enable_headless": True,
			"enable_scrapy": True,
			"force_update": False,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["cb_folderpath"] = self.DEFAULT_CB_FOLDERPATH if self.xcfg["cb_folderpath"] is None else self.xcfg["cb_folderpath"]
		self.xcfg["cb_data_folderpath"] = os.path.join(self.xcfg["cb_folderpath"], self.DEFAULT_CB_DATA_FOLDERNAME) if self.xcfg["cb_data_folderpath"] is None else self.xcfg["cb_data_folderpath"]

		self.can_scrape = self.__can_scrape()
		self.requests_module = None
		self.beautifulsoup_class = None
		self.web_driver = None
		self.selenium_select_module = None
		self.wget_module = None
		self.cb_publish_detail = {}

		self.cb_ignore_set = set()
		self.cb_stock_ignore_set = set()
# Don't consider the CB temporarily
		if self.xcfg["cb_ignore_list"] is not None:
			# import pdb; pdb.set_trace()
			if type(self.xcfg["cb_ignore_list"]) is str:
				self.xcfg["cb_ignore_list"] = self.xcfg["cb_ignore_list"].split(",")
			self.cb_ignore_set = set(self.xcfg["cb_ignore_list"])
			self.cb_stock_ignore_set = set(map(lambda x: x[:4], self.cb_ignore_set))

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


	def __enter__(self):
# # Open the workbook of cb quotation
# 		self.cb_workbook = xlrd.open_workbook(self.xcfg["cb_quotation_filepath"])
# 		self.cb_worksheet = self.cb_workbook.sheet_by_index(0)
# # Open the workbook of cb stock quotation
# 		self.cb_stock_workbook = xlrd.open_workbook(self.xcfg["cb_stock_quotation_filepath"])
# 		self.cb_stock_worksheet = self.cb_stock_workbook.sheet_by_index(0)
		return self


	def __exit__(self, type, msg, traceback):
		if self.web_driver is not None:
			self.web_driver.close()
			self.web_driver = None
		# if self.cb_workbook is not None:
		# 	self.cb_workbook.release_resources()
		# 	del self.cb_workbook
		# 	self.cb_workbook = None
		# if self.cb_stock_workbook is not None:
		# 	self.cb_stock_workbook.release_resources()
		# 	del self.cb_stock_workbook
		# 	self.cb_stock_workbook = None
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


	def __get_web_driver(self, web_driver_filepath=r"C:\chromedriver.exe"):
		if not self.__check_selenium_module_installed():
			raise RuntimeError("The selenium module is NOT installed!!!")
		if self.web_driver is None:
			module = __import__("selenium.webdriver")
			web_driver_class = getattr(module, "webdriver")
			# self.web_driver = web_driver_class.Chrome()
			options = None
			if self.xcfg["enable_headless"]:
				options = web_driver_class.EdgeOptions()
				options.add_argument("--headless")
# Valids levels are INFO = 0, WARNING = 1, LOG_ERROR = 2, LOG_FATAL = 3
				options.add_argument("--log-level=3")
			self.web_driver = web_driver_class.Edge(options=options)
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
		pattern = r"(.+)\(([\d]{5,6})\)"
		cb_data = {}
		with open(self.xcfg["cb_summary_filepath"], newline='') as f:
			rows = csv.reader(f)
			regex = re.compile(pattern)
			title_list = None
			for index, row in enumerate(rows):
				# import pdb; pdb.set_trace()
				if index in [0, 1,]: pass
				elif index == 2:
					title_list = list(map(lambda x: x.lstrip("=\"").rstrip("\"").rstrip("(%)"), row))
# ['可轉債商品', '可轉債收盤價', '成交量', '標的股收盤價', '轉換價值', '轉換溢價率', '轉換價格生效日', '轉換價格', '每張可轉換股數', '下一賣回日', '下一賣回價', '提前賣回收益率', '到期收益率', '發行日', '到期日', '發行張數', '剩餘張數', '轉換比例', '票面利率', '擔保情形', '']
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
				# print("%s" % (",".join(data_list)))
		# import pdb; pdb.set_trace()
		return cb_data


	def __print_file_modification_date(self):
		print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
		print("%s  修改時間: %s" % (os.path.basename(self.xcfg["cb_summary_filepath"]), self.xcfg["cb_summary_file_modification_date_str"]))
		print("%s  修改時間: %s" % (os.path.basename(self.xcfg["cb_list_filepath"]), self.xcfg["cb_list_file_modification_date_str"]))
		print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++\n")


	def __sort_future_publishing_cb_list(self, cb_publish_diff_quotation_id_list, shift_day=0):
		today = datetime.fromordinal(date.today().toordinal())
		if shift_day != 0:
# shift_day: (1) <0 => before today, (2) >0 => after today
			today += relativedelta(days=shift_day)
		cb_publish_diff_quotation_id_list = list(filter(lambda x: self.__strptime(self.cb_publish[x]['發行日期']) >= today, cb_publish_diff_quotation_id_list))  # 已全數轉換為普通股 資料尚未更新
		cb_publish_diff_quotation_id_list.sort(key=lambda x: self.__strptime(self.cb_publish[x]['發行日期']))
		return cb_publish_diff_quotation_id_list


	def __get_future_publishing_cb_list(self, shift_day=0):
		cb_publish_id_set = set(self.cb_publish.keys())  # - self.cb_ignore_set
		cb_quotation_id_set = set(self.cb_id_list)  # - self.cb_ignore_set
		cb_publish_diff_quotation_id_list = list(cb_publish_id_set - cb_quotation_id_set - self.cb_ignore_set) 
		return self.__sort_future_publishing_cb_list(cb_publish_diff_quotation_id_list, shift_day=shift_day) if (len(cb_publish_diff_quotation_id_list) > 0) else None


	def scrape_publish_cb(self):
		# import pdb; pdb.set_trace()
		driver = self.__get_web_driver()
		driver.get(self.PUBLISH_CB_LIST_URL)
		time.sleep(5)
		table = driver.find_element("xpath", '//*[@id="tables-content"]/div/div[2]/div[2]/table')
		trs = table.find_elements("tag name", "tr")
		for tr in trs:
			tds = tr.find_elements("tag name", "td")


	def calculate_internal_rate_of_return(self, cb_quotation, use_percentage=True):
		irr_dict = {}
		# import pdb; pdb.set_trace()
		adjusted_value = self.DEFAULT_CB_TRANSACTION_COST / 1000.0
		for cb_id in self.cb_id_list:
			cb_quotation_data = cb_quotation[cb_id]
			# print(cb_quotation_data)
			no_data = True
			days = None
			if cb_quotation_data["賣出一"] is not None:
				days = self.__get_days(cb_quotation_data["到期日"])
				if days != 0:
					no_data = False
			if no_data:
				irr_dict[cb_id] = {"商品": cb_quotation_data["商品"], "到期日": cb_quotation_data["到期日"], "賣出一": cb_quotation_data["賣出一"], "到期天數": None, "年化報酬率": None}
				# print("%s %s" % (cb_id, str(irr_dict)))
			else:
				days_to_year = days / 365.0
				# print("商品: %s, 到期日: %s, days: %d, days_to_year: %f" % (cb_quotation_data["商品"], cb_quotation_data["到期日"], days, days_to_year))
				irr = math.pow(100.0 / cb_quotation_data["賣出一"], 1 / days_to_year) - 1
				adjusted_sell_price = float(cb_quotation_data["賣出一"]) + adjusted_value
				adjusted_irr = math.pow(100.0 / adjusted_sell_price, 1 / days_to_year) - 1
				if use_percentage:
					irr *= 100.0
					adjusted_irr *= 100.0
				irr_dict[cb_id] = {"商品": cb_quotation_data["商品"], "到期日": cb_quotation_data["到期日"], "賣出一": cb_quotation_data["賣出一"], "到期天數": days, "年化報酬率": irr, "年化報酬率(扣手續費)": adjusted_irr}
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
				# print(cb_quotation_data)
				# print(cb_stock_quotation_data)
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
		def filter_func(x):
			return True if ((x[1]["股票溢價率"] is not None) and (abs(x[1]["股票溢價率"]) <= absolute_threshold)) else False
		stock_premium_dict = dict(filter(filter_func, stock_premium_dict.items()))
		if duration_within_days is not None:
			# # duration_date = datetime.now() + timedelta(days=duration_within_days)
			# # irr_dict = dict(filter(lambda x: datetime.strptime(x[1]["到期日"],"%Y/%m/%d") <= duration_date, irr_dict.items()))
			# stock_premium_dict = dict(filter(lambda x: x[1]["到期天數"] <= duration_within_days, stock_premium_dict.items()))
			def filter_func1(x):
				return True if ((x[1]["到期天數"] is not None) and (x[1]["到期天數"] <= duration_within_days)) else False
			stock_premium_dict = dict(filter(filter_func1, stock_premium_dict.items()))
		if need_sort:
			stock_premium_dict = collections.OrderedDict(sorted(stock_premium_dict.items(), key=lambda x: x[1]["股票溢價率"], reverse=False))
		return stock_premium_dict


	def search_cb_mass_volume(self, cb_quotation, cb_stock_quotation, volume_daily2total_ratio_percentage_threshold=3.5, volume_top_position_ratio_percentage_threshold=20, maturity_date_threshold=180, low_conversion_premium_rate_threshold=10, breakeven_threshold=110, use_percentage=True):
		# premium_dict = self.calculate_premium(cb_quotation, cb_stock_quotation, need_breakeven=True, use_percentage=True)
		no_filter = (low_conversion_premium_rate_threshold is None) and (breakeven_threshold is None)
		volume_daily2total_ratio_percentage_cb_dict = {}
		volume_top_position_ratio_cb_dict = {}
		cb_id_list_len = len(self.cb_id_list)
		volume_top_position_threshold = int(cb_id_list_len * volume_top_position_ratio_percentage_threshold / 100.0) 
		# import pdb; pdb.set_trace()
		for cb_index, cb_id in enumerate(self.cb_id_list):
			cb_data_dict = self.__collect_cb_full_data(cb_id, cb_quotation, cb_stock_quotation, use_percentage)
			if (not no_filter) and (cb_data_dict["成交"] is None):
				# print("Ignore CB[%s]: 沒有 成交" % cb_stock_id)
				continue
			if low_conversion_premium_rate_threshold is not None and cb_data_dict["溢價率"] is not None and cb_data_dict["溢價率"] > low_conversion_premium_rate_threshold:
				continue
			if breakeven_threshold is not None and cb_data_dict["成交"] > breakeven_threshold:
				continue
			# print("id: %s, index: %d, 總量: %d" % (cb_id, cb_index, cb_data_dict["總量"]))
			volume_daily2total_ratio = 100.0 * cb_data_dict["總量"] / cb_data_dict["發行張數"]
			if volume_daily2total_ratio >= volume_daily2total_ratio_percentage_threshold:
				volume_daily2total_ratio_percentage_cb_dict[cb_id] = copy.deepcopy(cb_data_dict)
				volume_daily2total_ratio_percentage_cb_dict[cb_id]["佔比"] = volume_daily2total_ratio
			if cb_index <= volume_top_position_threshold:
				volume_top_position_ratio_cb_dict[cb_id] = copy.deepcopy(cb_data_dict)
				volume_top_position_ratio_cb_dict[cb_id]["排名"] = cb_index
		return volume_daily2total_ratio_percentage_cb_dict, volume_top_position_ratio_cb_dict


	@retry()
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
		try:
			table_trs = table[0].find_all("tr")
		except Exception as e:
# Too many query requests from your ip, please wait and try again later!!
			print(e)
			raise RetryException("Too many query requests from your ip, please wait and try again later")
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

		table_row_start_index = None
		title_tds = None
		for index in range(0, 4):
			tds = trs[index].find_elements("tag name", "td")
			if len(tds) > 1:
				table_row_start_index = index
				title_tds = tds
				break
		if table_row_start_index is None:
			raise RuntimeError("Fails to parse the table in 季盈餘")
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
		if len(trs) <= 5:
			return None
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
		table = None
		try:
			table = driver.find_element("xpath", '//*[@id="oMainTable"]')
		# except selenium.common.exceptions.NoSuchElementException as e:
		except Exception as e:
			print("WARNING: No data in 主力進出, due to: %s" % str(e))
			return data_dict
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
		title_list_len = len(title_list)
		# import pdb; pdb.set_trace()
		table_row_start_index += 1
		# table_row_end_index = table_row_start_index + 15
		tr_cnt = 0
		for index, tr in enumerate(trs[table_row_start_index:]):
			tds = tr.find_elements("tag name", "td")
			td_text_list = []
			for td in tds:
				td_text_list.append(td.text)
			if len(td_text_list) != title_list_len:
				break
			tr_cnt += 1
			# import pdb; pdb.set_trace()
			if td_text_list[0] != ' ':
				data_dict[main_title_list[0]][tds[0].text] = dict(zip(title_list[1:5], td_text_list[1:5]))
			if td_text_list[5] != ' ':
				data_dict[main_title_list[1]][tds[5].text] = dict(zip(title_list[6:], td_text_list[6:]))
		table_row_start_index += tr_cnt
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
		daily_new_data_date_tmp = data_scrapy_date + timedelta(days=1)
		daily_new_data_date = datetime(daily_new_data_date_tmp.year, daily_new_data_date_tmp.month, daily_new_data_date_tmp.day)
		return daily_new_data_date


	def __calculate_stock_info_monthly_update_time(self, data_scrapy_date):
		monthly_new_data_date = None
		if data_scrapy_date.day < 10:
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
		data_scrapy_date = datetime.strptime(data_scrapy_timestr, "%Y/%m/%d %H:%M:%S")
		return (self.STOCK_INFO_UPDATE_TIME_FUNCPTR_DICT[update_frequency])(data_scrapy_date)
		# print("Now: %s\nDaily: %s\nMonthly: %s\nQuarterly: %s\nYearly: %s" % (data_scrapy_timestr, str(daily_new_data_date), str(monthly_new_data_date), str(quarterly_new_data_date), str(yearly_new_data_date)))


# Only evaluate that it's required to scrape data if dry_run is True
	def scrape_stock_info(self, cb_id, from_file=True, dry_run=False):
		data_dict = None
		# import pdb; pdb.set_trace()
		driver = None
		cb_stock_id = cb_id[:4]
		# import pdb; pdb.set_trace()
		data_filepath = os.path.join(self.xcfg["cb_data_folderpath"], "%s.txt" % cb_stock_id)
		if from_file:
			if self.__check_file_exist(data_filepath):
				if (not dry_run) and self.xcfg["force_update"]:
					try:
						os.remove(data_filepath)
					except OSError as e:
						print("Fails to delete the file: %s, due to: %s" % (data_filepath, str(e)))
				else:
					with open(data_filepath, "r", encoding='utf-8') as f:
# Load JSON data using OrderedDict as the object_pairs_hook. The type of the data_dict variable is OrderedDict
						'''
						while Python 3.7+ maintains the insertion order in regular dictionaries, this behavior is not part of the Python language 
						specification and could vary in different implementations or versions. If order preservation is crucial, using OrderedDict 
						or specifying an object_pairs_hook when loading JSON data is a more reliable approach
						'''
						data_dict = json.load(f, object_pairs_hook=OrderedDict)
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
			need_scrapy_count = 0
			for scrapy_key, scrapy_funcptr in self.STOCK_INFO_SCRAPY_FUNCPTR_DICT.items():
				# import pdb; pdb.set_trace()
# Check it's required to scrape the data
				need_scrapy = True
				if scrapy_key in data_time_dict:
					data_time_str = data_time_dict[scrapy_key]
					update_frequency = self.STOCK_INFO_SCRAPY_UPDATE_FREQUENCY_DICT[scrapy_key]
					new_update_time = self.calculate_stock_info_update_time(data_time_str, update_frequency)
					if new_update_time > cur_datetime:
						need_scrapy = False
				if need_scrapy:
					if dry_run:
						return True
					need_scrapy_count += 1
					if driver is None:
						driver = self.__get_web_driver()
					url = self.STOCK_INFO_SCRAPY_URL_FORMAT_DICT[scrapy_key] % cb_stock_id
					print("Scrape %s:%s......" % (cb_stock_id, scrapy_key))
					start_time = time.time()
					driver.get(url)
					time.sleep(5)
					scrapy_data = scrapy_funcptr(driver)
					if scrapy_data is None:
						print("WARNING: No data in %s:%s, URL: %s" % (cb_stock_id, scrapy_key, url))
					data_content_dict[scrapy_key] = scrapy_data
					end_time = time.time()
					print("Scrape %s:%s...... Done in %d seconds" % (cb_stock_id, scrapy_key, (end_time - start_time)))
					data_time_dict[scrapy_key] = cur_datetime.strftime("%Y/%m/%d %H:%M:%S")
			if dry_run:
				return False
			# print(data_dict)
			total_end_time = time.time()
			if need_scrapy_count > 0:
					print("Scrape All...... Done in %d seconds" % (total_end_time - total_start_time))
# Writing to file
			# import pdb; pdb.set_trace()
			if from_file and need_scrapy_count > 0:
				with open(data_filepath, "w", encoding='utf-8') as f:
					'''
					starting from Python 3.7, the keys of a dictionary are guaranteed to be in the same order as they were inserted 
					when using the built-in dict type. This means that when you call json.dump to write a dictionary to a JSON file in 
					Python 3.7 and later, the order of keys will be preserved in the resulting JSON file.
					'''
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
			mobj = re.search(r".+([\d]{5})", span.text)
			if mobj is None:
				raise RuntimeError("Fail to find the month of the table")
			table_month = mobj.group(1)
			filename = self.DEFAULT_CB_MONTHLY_CONVERT_DATA_FILENAME_PREFIX + table_month
			filepath = os.path.join(self.xcfg["cb_data_folderpath"], filename)
			# import pdb; pdb.set_trace()
			cur_datetime = datetime.now()
			need_scrapy = True
			if self.__check_file_exist(filepath):
				# print("The file[%s] already exist !!!" % filepath)
				with open(filepath, "r", encoding='utf8') as f:
					data_dict = json.load(f)
					# Check it's required to scrape the data				
					data_time_str = data_dict["time"]
					new_update_time = self.calculate_stock_info_update_time(data_dict["time"], "Monthly")
					if new_update_time > cur_datetime:
						need_scrapy = False
			if need_scrapy:
				data_dict = {
					"time": None, 
					"content": {},
				}
# thead
				print("The file[%s] does NOT exist. Scrape the data from website" % filepath)
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
					data_dict["content"][td_text_list[0]] = dict(zip(table_title_list[1:], td_text_list[1:]))
				data_dict["time"] = cur_datetime.strftime("%Y/%m/%d %H:%M:%S")
				# time.sleep(5)
				# import pdb; pdb.set_trace()
# Writing to file
				with open(filepath, "w", encoding='utf-8') as f:
					json.dump(data_dict, f, indent=3, ensure_ascii=False)
		except Exception as e:
			print("Exception occurs while scraping [%s], due to: %s" % (url, str(e)))
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
			

	def search_cb_mass_convert(self, cb_quotation=None, cb_stock_quotation=None, table_month=None, mass_convert_threshold=-10.0, filter_cb=True):
		# import pdb; pdb.set_trace()
		mass_convert_cb_dict = None
		def filter_funcptr(x):
			# import pdb; pdb.set_trace()
			# print(x)
			data_dict = x[1]
			if int(data_dict["發行張數"]) == 0: return 0
			# print(float(data_dict["增減數額"]) / float(data_dict["發行張數"]))
			return float(data_dict["增減數額"]) / float(data_dict["發行張數"]) * 100.0
		def update_funcptr(x):
			# import pdb; pdb.set_trace()
			cb_data = None
			try:
				cb_data = self.__collect_cb_full_data(x[0], cb_quotation, cb_stock_quotation)
				cb_data.update(x[1])
			except ValueError as e:
				print("Error: %s" % str(e))
			# print(cb_data)
			return cb_data
		try:
			cb_monthly_convert_data = self.get_cb_monthly_convert_data(table_month)
			# import pdb; pdb.set_trace()
			convert_cb_dict = cb_monthly_convert_data["content"]
			if filter_cb:
				if not self.xcfg['cb_all']:
					convert_cb_dict = dict(filter(lambda x: x[0] in self.cb_id_list, convert_cb_dict.items()))
			# mass_convert_cb_dict = dict(filter(lambda x: float(x[1]["增減百分比"]) < mass_convert_threshold, convert_cb_dict.items()))
			mass_convert_cb_dict = dict(filter(lambda x: filter_funcptr(x) < mass_convert_threshold, convert_cb_dict.items()))
			if cb_quotation is not None and cb_stock_quotation is not None:
				mass_convert_cb_dict_tmp = mass_convert_cb_dict
				mass_convert_cb_dict = dict([(x[0], update_funcptr(x)) for x in mass_convert_cb_dict_tmp.items()])
		except ValueError as e:
			print("CB Mass Convert: %s" % str(e))
			return None
		# import pdb; pdb.set_trace()
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
		# # print(data_dict_summary)
		quotation_data_dict = self.__read_cb_quotation()
		stock_quotation_data_dict = self.__read_cb_stock_quotation()
		if self.xcfg['cb_all']:
			self.check_data_source(quotation_data_dict, stock_quotation_data_dict)
		print("\n*****************************************************************\n")

		# print(quotation_data_dict)
		# print(self.calculate_internal_rate_of_return(quotation_data_dict))
		irr_dict = self.get_positive_internal_rate_of_return(quotation_data_dict)
		if bool(irr_dict):
			print("=== 年化報酬率 ==================================================")
			title_list = ["年化報酬率", "年化報酬率(扣手續費)", "賣出一", "到期日",]
			print("  ===> %s" % ", ".join(title_list))
			for irr_key, irr_data in irr_dict.items():
				print("%s[%s]: %.2f  %.2f  %.2f  %s" % (irr_data["商品"], irr_key, float(irr_data["年化報酬率"]), float(irr_data["年化報酬率(扣手續費)"]), float(irr_data["賣出一"]), irr_data["到期日"]))
			print("=================================================================\n")
		premium_dict = self.get_negative_premium(quotation_data_dict, stock_quotation_data_dict)
		if bool(premium_dict):
			print("=== 溢價率(套利) ================================================")
			title_list = ["溢價率", "融資餘額", "融券餘額",]
			print("  ===> %s" % ", ".join(title_list))
			# import pdb; pdb.set_trace()
			for premium_key, premium_data in premium_dict.items():
				print("%s[%s]: %.2f  %d  %d" % (premium_data["商品"], premium_key, float(premium_data["溢價率"]), premium_data["融資餘額"], premium_data["融券餘額"]))
			print("=================================================================\n")
		stock_premium_dict = self.get_absolute_stock_premium(quotation_data_dict, stock_quotation_data_dict)
		if bool(stock_premium_dict):
			print("=== 股票溢價率 ==================================================")
			title_list = ["股票溢價率",]
			print("  ===> %s" % ", ".join(title_list))
			for stock_premium_key, stock_premium_data in stock_premium_dict.items():
				print("%s[%s]: %.2f" % (stock_premium_data["商品"], stock_premium_key, float(stock_premium_data["股票溢價率"])))
			print("=================================================================\n")
		cb_dict = self.get_low_premium_and_breakeven(quotation_data_dict, stock_quotation_data_dict)
		if bool(cb_dict):
			print("=== 低溢價且保本 ================================================")
			title_list = ["溢價率", "成交", "賣出一", "到期日",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in cb_dict.items():
				print("%s[%s]: %.2f  %.2f  %.2f  %s" % (cb_data["商品"], cb_key, float(cb_data["溢價率"]), float(cb_data["成交"]), float(cb_data["賣出一"]), cb_data["到期日"]))
			print("=================================================================\n")
		volume_daily2total_ratio_percentage_cb_dict, volume_top_position_ratio_cb_dict = self.search_cb_mass_volume(quotation_data_dict, stock_quotation_data_dict)
		if bool(volume_daily2total_ratio_percentage_cb_dict):
			print("=== 當日成交量對發行總量比例大 ==================================================")
			title_list = ["日期", "溢價率", "成交", "總量", "發行張數", "佔比", '一週%', '一月%', '一季%', "成交(股)",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in volume_daily2total_ratio_percentage_cb_dict.items():
				# print("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
				print("%s[%s]:  %s  %s  %s  %s  %s  %s  %s  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["發行日期"], self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), self.__float(cb_data["佔比"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"], self.__float(cb_data["成交(股)"])))
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
		if bool(volume_top_position_ratio_cb_dict):
			print("=== 當日成交量大 ==================================================")
			title_list = ["日期", "溢價率", "成交", "總量", "發行張數", "排名", '一週%', '一月%', '一季%', "成交(股)",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in volume_top_position_ratio_cb_dict.items():
				# print("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
				print("%s[%s]:  %s  %s  %s  %s  %s  %s  %s  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["發行日期"], self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), self.__int(cb_data["排名"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"], self.__float(cb_data["成交(股)"])))
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

		issuing_date_cb_dict, convertible_date_cb_dict, maturity_date_cb_dict = self.search_cb_opportunity_dates(quotation_data_dict, stock_quotation_data_dict)
		# issuing_date_cb_dict, convertible_date_cb_dict, maturity_date_cb_dict = self.search_cb_opportunity_dates(quotation_data_dict, stock_quotation_data_dict, low_conversion_premium_rate_threshold=None, breakeven_threshold=None)
		if bool(issuing_date_cb_dict):
			print("=== 近發行日期 ==================================================")
			title_list = ["日期", "天數", "溢價率", "成交", "總量", "發行張數", '一週%', '一月%', '一季%',]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in issuing_date_cb_dict.items():
				# print("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
				print("%s[%s]:  %s(%d)  %s  %s  %s  %s  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
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
			title_list = ["日期", "天數", "溢價率", "成交", "總量", "發行張數", '一週%', '一月%', '一季%',]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in convertible_date_cb_dict.items():
				# print("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
				print("%s[%s]:  %s(%d)  %s  %s  %s  %s  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_data["發行張數"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
				cb_publish_detail_dict = self.get_publish_detail(cb_key)
				print(" *************")
				print("  本月受理轉(交)換之公司債張數: %s" % (cb_publish_detail_dict["本月受理轉(交)換之公司債張數"]))
				print("  最新轉(交)換價格: %s" % (cb_publish_detail_dict["最新轉(交)換價格"]))
				print("  最近轉(交)換價格生效日期: %s" % (cb_publish_detail_dict["最近轉(交)換價格生效日期"]))
				print(" *************")
			print("=================================================================\n")
		if bool(maturity_date_cb_dict):
			print("=== 近到期日期 ==================================================")
			title_list = ["日期", "天數", "溢價率", "成交", "總量", "發行總面額(億)", "股本(億)", '一週%', '一月%', '一季%',]
			print("  ===> %s" % ", ".join(title_list))
			cb_monthly_convert_data = self.get_cb_monthly_convert_data()
			for cb_key, cb_data in maturity_date_cb_dict.items():
				cb_stock_key = cb_key[:4]
				cb_stock_data = stock_quotation_data_dict[cb_stock_key]
				cb_publish_data = self.cb_publish[cb_key]
				mass_convert_exist = True if (cb_key in cb_monthly_convert_data["content"]) else False
				mass_convert_cb_data = cb_monthly_convert_data["content"][cb_key] if mass_convert_exist else None
				# import pdb; pdb.set_trace()
				# print("%s[%s]:  %s(%d)  %.2f  %.2f  %d  %d  %.2f  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(mass_convert_cb_data["發行張數"]) * 100000 / 100000000, self.__float(cb_stock_data["股本"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
				print("%s[%s]:  %s(%d)  %s  %s  %s  %s  %s  %s  %s  %s" % (cb_data["商品"], cb_key, cb_data["日期"], self.__int(cb_data["天數"]), self.__float(cb_data["溢價率"]), self.__float(cb_data["成交"]), self.__int(cb_data["總量"]), self.__int(cb_publish_data["發行總面額"]) / 100000, self.__float(cb_stock_data["股本"]), cb_data["一週%"], cb_data["一月%"], cb_data["一季%"]))
				cb_publish_detail_dict = self.get_publish_detail(cb_key)
				print(" *************")
				print("  本月受理轉(交)換之公司債張數: %s" % (cb_publish_detail_dict["本月受理轉(交)換之公司債張數"]))
				print("  最新轉(交)換價格: %s" % (cb_publish_detail_dict["最新轉(交)換價格"]))
				print("  股票成交價格: %s" % (stock_quotation_data_dict[cb_stock_key]["成交"]))
				if mass_convert_exist:
					print("  發行張數: %d" % int(mass_convert_cb_data["發行張數"]))
					print("  前月底保管張數: %d" % int(mass_convert_cb_data["前月底保管張數"]))
					print("  本月底保管張數: %d" % int(mass_convert_cb_data["本月底保管張數"]))
					no_convert_percentage = float(mass_convert_cb_data["本月底保管張數"])/int(cb_data["發行張數"]) * 100.0
					print("  未轉換百分比: %.2f" % no_convert_percentage)
					monthly_convert_percentage = float(mass_convert_cb_data["增減數額"]) / float(mass_convert_cb_data["發行張數"]) * 100.0
					print("  增減百分比: %.2f" % monthly_convert_percentage)
					# print("  最近轉(交)換價格生效日期: %s" % (cb_publish_detail_dict["最近轉(交)換價格生效日期"]))
				print(" *************")
			print("=================================================================\n")
		mass_convert_cb_dict = self.search_cb_mass_convert(quotation_data_dict, stock_quotation_data_dict)
		if bool(mass_convert_cb_dict):
			print("=== CB大量轉換 ==================================================")
			title_list = ["增減百分比", "前月底保管張數", "本月底保管張數", "發行張數", "到期日期",]
			print("  ===> %s" % ", ".join(title_list))
			for cb_key, cb_data in mass_convert_cb_dict.items():
				if cb_data is None: 
					print("%s:  missing data" % cb_key)
				else:
					mass_convert_percentage = float(cb_data["增減數額"]) / float(cb_data["發行張數"]) * 100.0
					print("%s[%s]:  %.2f  %d  %d  %d  %s" % (cb_data["名稱"], cb_key, mass_convert_percentage, int(cb_data["前月底保管張數"]), int(cb_data["本月底保管張數"]), int(cb_data["發行張數"]), cb_data["到期日期"]))
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


	def print_tracked_cb(self):
		if self.xcfg["cb_list"] is None:
			self.__get_cb_list_from_file()
		for cb in self.xcfg["cb_list"]:
			print(cb)


	def modify_tracked_cb(self, modify_cb_list_str):
		# import pdb; pdb.set_trace()
		if self.xcfg["cb_list"] is None:
			self.__get_cb_list_from_file()
		modify_cb_list = modify_cb_list_str.split(",")
		for modify_cb in modify_cb_list:
			if modify_cb[0] == "+":
				add_cb = modify_cb[1:]
				if add_cb in self.xcfg["cb_list"]:
					print("The CB[%s] already exists in the list" % add_cb)
				else:
					self.xcfg["cb_list"].append(add_cb)
			elif modify_cb[0] == "x":
				remove_cb = modify_cb[1:]
				if remove_cb not in self.xcfg["cb_list"]:
					print("The CB[%s] does NOT exist in the list" % remove_cb)
				else:
					self.xcfg["cb_list"].remove(remove_cb)
			else:
				raise ValueError("Incorrect operator: %s" % modify_cb)
		# import pdb; pdb.set_trace()
		self.xcfg["cb_list"] = list(filter(lambda x: len(x) != 0, self.xcfg["cb_list"]))
		with open(self.xcfg['cb_list_filepath'], 'w') as fp:
			for line in self.xcfg["cb_list"]:
				fp.write("%s\n" % line)


	def track(self, data_days=5):
# ['商品', '成交', '漲幅%', '總量', '買進一', '賣出一', '到期日']
		quotation_data_dict = self.__read_cb_quotation()
		stock_quotation_data_dict = self.__read_cb_stock_quotation()
		if self.xcfg['cb_all']:
			self.check_data_source(quotation_data_dict, stock_quotation_data_dict)
		mass_convert_cb_dict = self.search_cb_mass_convert(quotation_data_dict, stock_quotation_data_dict, filter_cb=False)

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
			# import pdb; pdb.set_trace()
			scrapy_data = self.scrape_stock_info(cb_stock_id) if self.xcfg["enable_scrapy"] else None
			print("%s[%s]:" % (quotation_data["商品"], cb_id))
			print("====================================================================")
			print(" %s" % "  ".join(["溢價率", "成交", "賣出一", "轉換價格", "成交(股)",]))
			try:
				print(" %.2f  %.2f  %.2f  %.2f  %.2f" % (float(premium_data["溢價率"]), float(quotation_data["成交"]), float(quotation_data["賣出一"]), float(summary_data["轉換價格"]), float(stock_quotation_data["成交"])))
			except TypeError:
				print(" %s" % ("  ".join([str(premium_data["溢價率"]), str(quotation_data["成交"]), str(quotation_data["賣出一"])])))
			print(" %s" % "  ".join(["發行日期", "年期", "發行總張數", "發行總面額(億)", "股本(億)",]))
			try:
				print(" %s  %s  %d  %d  %.2f" % (publish_data["發行日期"], publish_data["年期"], int(publish_data["發行總面額"]) / 100000, int(publish_data["發行總面額"]) / 100000000, stock_quotation_data["股本"]))
			except TypeError:
				print(" %s" % ("  ".join([str(premium_data["溢價率"]), str(quotation_data["成交"]), str(quotation_data["賣出一"])])))
# Check mass convert
			# import pdb; pdb.set_trace()
			mass_convert_cb_list = list(filter(lambda x: x[:4] == cb_stock_id, mass_convert_cb_dict.keys()))
			if len(mass_convert_cb_list) != 0:
				print("=== CB大量轉換 ==================================================")
				title_list = ["增減百分比", "前月底保管張數", "本月底保管張數", "發行張數", "到期日期"]
				for cb_id in mass_convert_cb_list:
					cb_data = mass_convert_cb_dict[cb_id]
					mass_convert_percentage = float(cb_data["增減數額"]) / float(cb_data["發行張數"]) * 100.0
					print(" %s  增減百分比: %.2f  前月底保管張數: %d, 本月底保管張數: %d, 發行張數: %d, 到期日期: %s" % (cb_data["名稱"], mass_convert_percentage, int(cb_data["前月底保管張數"]), int(cb_data["本月底保管張數"]), int(cb_data["發行張數"]), cb_data["到期日期"]))
# Data from Scrapy
			if not self.xcfg["enable_scrapy"]: continue
			dt_now = datetime.now()

			print("\n籌碼面")
			data_dicts = list(scrapy_data["content"]["法人持股"].items())
			data_dicts_len = len(data_dicts)
			data_dicts_count = min(data_days, data_dicts_len)
			assert data_dicts_count >= 2, "data_dicts_count[法人持股:%d] should be greater than 2" % data_dicts_count
			dt_data_date = self.__strptime(data_dicts[0][0], year_transform=True)
			dt_data = datetime(dt_data_date.year, dt_data_date.month, dt_data_date.day, 22)
			date_shift = 1 if dt_now < dt_data else 0
			print(" %s" % " ".join(["日期", "外資持股比重", "三大法人持股比重",]))
			for data_dict in data_dicts[date_shift:min(data_dicts_count + date_shift, data_dicts_len)]:
				print(" %s" % "  ".join([data_dict[0], data_dict[1]["外資持股比重"], data_dict[1]["三大法人持股比重"]]))
			# data_dicts = list(scrapy_data["content"]["融資融券"].items())
			# data_dicts_len = len(data_dicts)
			# data_dicts_count = min(data_days, data_dicts_len)
			# assert data_dicts_count >= 2, "data_dicts_count[融資融券:%d] should be greater than 2" % data_dicts_count
			# dt_data_date = self.__strptime(data_dicts[0][0], year_transform=True)
			# dt_data = datetime(dt_data_date.year, dt_data_date.month, dt_data_date.day, 22)
			# date_shift = 1 if dt_now < dt_data else 0
			# print(" %s" % " ".join(["日期", "融資餘額",]))
			# for data_dict in data_dicts[date_shift:min(data_dicts_count + date_shift, data_dicts_len)]:
			# 	print(" %s" % " ".join([data_dict[0], data_dict[1]["融資餘額"]]))

			print("\n基本面")
# Monthly
			# import pdb; pdb.set_trace()
			latest_data_dict = list(scrapy_data["content"]["月營收"].items())[0]
			print(" %s  營收年增率: %s" % (latest_data_dict[0], latest_data_dict[1]["年增率"]))
# Quarterly
			latest_data_dict = list(scrapy_data["content"]["獲利能力"].items())[0]
			print(" %s  EPS(元): %s" % (latest_data_dict[0], latest_data_dict[1]["EPS(元)"]))
			latest_data_dict = list(scrapy_data["content"]["季盈餘"].items())[0]
			# import pdb; pdb.set_trace()
			print(" %s  EPS年增率: %s" % (latest_data_dict[0], latest_data_dict[1]["年增率"]))
			latest_data_dict = list(scrapy_data["content"]["財務比率簡表(季)"].items())[0]
			if "營業毛利率" in latest_data_dict[1]:
				print(" %s  營業毛利率: %s  營業利益率: %s, 稅後淨利率: %s" % (latest_data_dict[0], latest_data_dict[1]["營業毛利率"], latest_data_dict[1]["營業利益率"], latest_data_dict[1]["稅後淨利率"]))
				print(" %s  營收成長率: %s  毛利成長率: %s, 營業利益成長率: %s, 稅後淨利成長率: %s" % (latest_data_dict[0], latest_data_dict[1]["營收成長率"], latest_data_dict[1]["營業毛利成長率"], latest_data_dict[1]["營業利益成長率"], latest_data_dict[1]["稅後淨利成長率"]))
				print(" %s  流動比率: %s  速動比率: %s, 負債比率: %s" % (latest_data_dict[0], latest_data_dict[1]["流動比率"], latest_data_dict[1]["速動比率"], latest_data_dict[1]["負債比率％"]))
			else:  # 銀行類股
				print(" %s  稅後淨利率: %s" % (latest_data_dict[0], latest_data_dict[1]["稅後淨利率(A)"]))
				print(" %s  營收成長率: %s  稅後淨利成長率: %s" % (latest_data_dict[0], latest_data_dict[1]["營收成長率(A)"], latest_data_dict[1]["稅後淨利成長率"]))
				print(" %s  負債比率: %s" % (latest_data_dict[0], latest_data_dict[1]["負債比率％"]))
			latest_data_dict = list(scrapy_data["content"]["資產負債簡表(季)"].items())[0]
			if "流動資產" in latest_data_dict[1]:
				print(" %s  流動資產: %s  流動負債: %s  股東權益總額: %s" % (latest_data_dict[0], latest_data_dict[1]["流動資產"], latest_data_dict[1]["流動負債"], latest_data_dict[1]["股東權益總額"]))
			else:  # 銀行類股
				print(" %s  股東權益總額: %s" % (latest_data_dict[0], latest_data_dict[1]["股東權益總額"]))
			latest_data_dict = list(scrapy_data["content"]["現金流量簡表(季)"].items())[0]
			value1 = int(latest_data_dict[1]["來自營運之現金流量"].replace(",", ""))
			value2 = int(latest_data_dict[1]["投資活動之現金流量"].replace(",", ""))
			free_cash_flow = "{:,}".format((value1 + value2))
			print(" %s  營運現金流: %s  投資現金流: %s, 融資現金流: %s, 自由現金流: %s" % (latest_data_dict[0], latest_data_dict[1]["來自營運之現金流量"], latest_data_dict[1]["投資活動之現金流量"], latest_data_dict[1]["融資活動之現金流量"], free_cash_flow))
# Yearly
			latest_data_dict = list(scrapy_data["content"]["財務比率簡表(年)"].items())[0]
			if "營業毛利率" in latest_data_dict[1]:
				print(" %s  營業毛利率: %s  營業利益率: %s, 稅後淨利率: %s" % (latest_data_dict[0], latest_data_dict[1]["營業毛利率"], latest_data_dict[1]["營業利益率"], latest_data_dict[1]["稅後淨利率"]))
				print(" %s  營收成長率: %s  毛利成長率: %s, 營業利益成長率: %s, 稅後淨利成長率: %s" % (latest_data_dict[0], latest_data_dict[1]["營收成長率"], latest_data_dict[1]["營業毛利成長率"], latest_data_dict[1]["營業利益成長率"], latest_data_dict[1]["稅後淨利成長率"]))
				print(" %s  流動比率: %s  速動比率: %s, 負債比率: %s" % (latest_data_dict[0], latest_data_dict[1]["流動比率"], latest_data_dict[1]["速動比率"], latest_data_dict[1]["負債比率％"]))
			else:  # 銀行類股
				print(" %s  稅後淨利率: %s" % (latest_data_dict[0], latest_data_dict[1]["稅後淨利率(A)"]))
				print(" %s  營收成長率: %s  稅後淨利成長率: %s" % (latest_data_dict[0], latest_data_dict[1]["營收成長率(A)"], latest_data_dict[1]["稅後淨利成長率"]))
				print(" %s  負債比率: %s" % (latest_data_dict[0], latest_data_dict[1]["負債比率％"]))
			latest_data_dict = list(scrapy_data["content"]["資產負債簡表(年)"].items())[0]
			if "流動資產" in latest_data_dict[1]:
				print(" %s  流動資產: %s  流動負債: %s  股東權益總額: %s" % (latest_data_dict[0], latest_data_dict[1]["流動資產"], latest_data_dict[1]["流動負債"], latest_data_dict[1]["股東權益總額"]))
			else:  # 銀行類股
				print(" %s  股東權益總額: %s" % (latest_data_dict[0], latest_data_dict[1]["股東權益總額"]))
			latest_data_dict = list(scrapy_data["content"]["現金流量簡表(年)"].items())[0]
			value1 = int(latest_data_dict[1]["來自營運之現金流量"].replace(",", ""))
			value2 = int(latest_data_dict[1]["投資活動之現金流量"].replace(",", ""))
			free_cash_flow = "{:,}".format((value1 + value2))
			print(" %s  營運現金流: %s  投資現金流: %s, 融資現金流: %s, 自由現金流: %s" % (latest_data_dict[0], latest_data_dict[1]["來自營運之現金流量"], latest_data_dict[1]["投資活動之現金流量"], latest_data_dict[1]["融資活動之現金流量"], free_cash_flow))


# Only evaluate that it's required to scrape data if dry_run is True
	def scrape_stock(self, stock_id_list, dry_run=False):
		if isinstance(stock_id_list, str):
			stock_id_list = stock_id_list.split(",")
		need_scrapy_stock_id_list = []			
		for stock_id in stock_id_list:
			ret = self.scrape_stock_info(stock_id, dry_run=dry_run)
			if dry_run and ret:
				need_scrapy_stock_id_list.append(stock_id)
		if dry_run:
			if len(need_scrapy_stock_id_list) == 0:
				print("All data are the latest")
				return False
			else:
				print("Some data are required to update: %s" % ", ".join(need_scrapy_stock_id_list))
				return True
		return None


	def recursive_scrape_stock(self, stock_id_list, retry_times=5):
		retry_count = 0
		while retry_count < retry_times:
			retry_count += 1
			# import pdb; pdb.set_trace()
			if not self.scrape_stock(stock_id_list, dry_run=True):
				return
			print("***** Scrape Stock...... %d *****" % retry_count)
			self.scrape_stock(stock_id_list)
		if not self.scrape_stock(stock_id_list, dry_run=True):
			raise RuntimeError("Fails to scrape all stock...")


# Only evaluate that it's required to scrape data if dry_run is True
	def scrape_stock_from_file(self, dry_run=False):
		stock_id_list = list(map(lambda x: x[:4], self.cb_id_list))
		return self.scrape_stock(stock_id_list, dry_run=dry_run)


	def recursive_scrape_stock_from_file(self, retry_times=5):
		stock_id_list = list(map(lambda x: x[:4], self.cb_id_list))
		# import pdb; pdb.set_trace()
		self.recursive_scrape_stock(stock_id_list, retry_times=retry_times)


# Only evaluate that it's required to scrape data if dry_run is True
	def scrape_stock_from_future_publishing_cb(self, dry_run=False):
		cb_publish_diff_quotation_id_list = self.__get_future_publishing_cb_list()
		if len(cb_publish_diff_quotation_id_list) != 0:
			new_public_cb_str_list = ["%s[%s]:  %s  %s年  %d張" % (self.cb_publish[cb_publish_id]['債券簡稱'], cb_publish_id, self.cb_publish[cb_publish_id]['發行日期'], self.cb_publish[cb_publish_id]['年期'], self.cb_publish[cb_publish_id]['發行總面額']/100000) for cb_publish_id in cb_publish_diff_quotation_id_list]
			print("=== 新發行 ===")
			for new_public_cb_str in new_public_cb_str_list: print(new_public_cb_str)
			self.scrape_stock(cb_publish_diff_quotation_id_list, dry_run=dry_run)
		else:
			print("=== * 暫無新發行 * ===")


	def recursive_scrape_stock_from_future_publishing_cb(self, retry_times=5):
		cb_publish_diff_quotation_id_list = self.__get_future_publishing_cb_list()
		if len(cb_publish_diff_quotation_id_list) != 0:
			new_public_cb_str_list = ["%s[%s]:  %s  %s年  %d張" % (self.cb_publish[cb_publish_id]['債券簡稱'], cb_publish_id, self.cb_publish[cb_publish_id]['發行日期'], self.cb_publish[cb_publish_id]['年期'], self.cb_publish[cb_publish_id]['發行總面額']/100000) for cb_publish_id in cb_publish_diff_quotation_id_list]
			print("=== 新發行 ===")
			for new_public_cb_str in new_public_cb_str_list: print(new_public_cb_str)
			self.recursive_scrape_stock(cb_publish_diff_quotation_id_list, retry_times=retry_times)
		else:
			print("=== * 暫無新發行 * ===")


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
	parser.add_argument('-l', '--list', required=False, action='store_true', help='List the potential targets based on the search rule.')
	parser.add_argument('-s', '--search', required=False, action='store_true', help='Select targets based on the search rule.')
	parser.add_argument('-t', '--track', required=False, action='store_true', help='Track specific targets.')
	parser.add_argument('-v', '--validate', required=False, action='store_true', help='Validate the estimation.')
	parser.add_argument('--cb_list', required=False, help='The list of specific CB targets.')
	parser.add_argument('--cb_ignore_list', required=False, help='The list of specific CB targets which are ignored.')
	parser.add_argument('--scrape_stock', required=False, help='Scrape the stock info of specific stocks and exit.')
	parser.add_argument('--scrape_stock_from_file', required=False, action='store_true', help="Scrape the stock info and exit. The scrapy stocks are from the 'cb_list' file")
	parser.add_argument('--scrape_stock_from_future_publishing_cb', required=False, action='store_true', help="Scrape the stock info of future publishing CB and exit")
	parser.add_argument('--recursive', required=False, action='store_true', help='Scrape stock recursively. Caution: Only take effect for the "scrape_stock", "scrape_stock_from_file" and "scrape_stock_from_future_publishing_cb" argument')
	parser.add_argument('--check_scrape_stock', required=False, help='Only check if it is required to scrape the stock info of specific stocks and exit.')
	parser.add_argument('--check_scrape_stock_from_file', required=False, action='store_true', help="Only check if it is required to scrape the stock info and exit. The scrapy stocks are from the 'cb_list' file")
	parser.add_argument('--check_scrape_stock_from_future_publishing_cb', required=False, action='store_true', help="Only check if it is required to scrape the stock info of future publishing CB and exit")
	parser.add_argument('--print_filepath', required=False, action='store_true', help='Print the filepaths used in the process and exit.')
	parser.add_argument('--print_tracked_cb', required=False, action='store_true', help='Print the CB list tracked in the file and exit.')
	parser.add_argument('--modify_tracked_cb', required=False, help='The rule for selecting targets. Default: 0.')
	parser.add_argument('--disable_headless', required=False, action='store_true', help='Disable headless web scrapy')
	parser.add_argument('--disable_scrapy', required=False, action='store_true', help='Disable data from scrapy. Caution: Only take effect for the "track" argument')
	parser.add_argument('--force_update', required=False, action='store_true', help='Force to scrape all data')
	args = parser.parse_args()

	cfg = {
		# "cb_list": "62822,62791,61965,54342,33881",
		# "cb_ignore_list", "",
	}
	if args.all:
		cfg['cb_all'] = args.all
	if args.cb_list:
		cfg['cb_list'] = args.cb_list
		if 'cb_all' in cfg:
			print("The 'all' flag is ignored...")
			cfg['cb_all'] = False
	if args.cb_ignore_list:
		cfg['cb_ignore_list'] = args.cb_ignore_list
	if args.disable_headless:
		cfg['enable_headless'] = False
	if args.disable_scrapy:
		cfg['enable_scrapy'] = False
	if args.force_update:
		cfg['force_update'] = True
	with ConvertibleBondAnalysis(cfg) as obj:
		obj.scrape_publish_cb()

