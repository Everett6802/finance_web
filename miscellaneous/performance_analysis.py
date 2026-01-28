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
from datetime import datetime, date, timedelta
import getpass
from collections import OrderedDict


class PerformanceAnalysis(object):

	DEFAULT_HOST_DATA_FOLDERPATH =  "C:\\Users\\%s\\project_data\\finance_web" % getpass.getuser()
	DEFAULT_DATA_FOLDERPATH =  os.getenv("DATA_PATH", DEFAULT_HOST_DATA_FOLDERPATH)
	DEFAULT_SOURCE_FILENAME = "加權指數歷史資料2000-2025.xlsx"
	DEFAULT_TIME_FIELD_NAME = "時間"	
	DEFAULT_CLOSING_PRICE_FIELD_NAME = "收盤價"	
	# DEFAULT_CONFIG_FOLDERPATH =  "C:\\Users\\%s" % os.getlogin()
	DEFAULT_DATE_BASE_NUMBER = 36526
	DEFAULT_DATE_BASE = date(2000, 1, 1)

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
	def mean(cls, data):
		return sum(data) / len(data)


	@classmethod
	def std(cls, data):
		mu = cls.mean(data)
		var = sum((x - mu) ** 2 for x in data) / len(data)
		return var ** 0.5


# daily_return就是每日的漲跌百分比 只是用「小數」表示，而不是用 %
# Cumulative Return: 2.31 => 這代表：Final=3.31×Initial 也就是 +231%（約 3.3 倍）
	@classmethod
	def cumulative_return(cls, daily_returns):
		value = 1.0
		for r in daily_returns:
			value *= (1 + r)
		return value - 1


	@classmethod
	def cumulative_return_from_prices(cls, prices):
		"""
		prices: list of prices (e.g. daily close)
		return: cumulative return (float)
		"""
		if len(prices) < 2:
			return 0.0
		return prices[-1] / prices[0] - 1


# cagr 和 cagr_from_prices_with_dates 分別用日曆年/交易年的差異 
# cagr傳進去的daily_returns 都是交易日 
# cagr_from_prices_with_dates算頭尾的時間間隔 有包含非交易日
# cumulative_return	「這段期間總共漲跌多少？」
# CAGR	「如果平均攤成年化，每年大概幾 %？」
	@classmethod
	def cagr(cls, daily_returns, periods_per_year=252):
		# import pdb; pdb.set_trace()
		total = cls.cumulative_return(daily_returns) + 1
		years = len(daily_returns) / periods_per_year
		return total ** (1 / years) - 1


	@classmethod
	def cagr_from_prices_with_dates(cls, prices, dates):
		"""
		prices: list of prices
		dates: list of date objects (datetime.date)
		"""
		if len(prices) < 2:
			return 0.0
		periods_per_year = 365
		start_price = prices[0]
		end_price = prices[-1]
		days = (dates[-1] - dates[0]).days
		years = days / periods_per_year
		if years <= 0:
			return 0.0
		return (end_price / start_price) ** (1 / years) - 1


# 把「每日波動度」轉換成「年化波動度」
# 假設：每日波動度 = 1% = 0.01  =>  年化波動度 = 0.01 * sqrt(252) = 0.1587 = 15.87%
	@classmethod
	def annualized_volatility(cls, daily_returns, periods_per_year=252):
		return cls.std(daily_returns) * (periods_per_year ** 0.5)


	@classmethod
	def sharpe_ratio(cls, daily_returns, risk_free_rate=0.0, periods_per_year=252):
		ann_return = cls.cagr(daily_returns, periods_per_year)
		ann_vol = cls.annualized_volatility(daily_returns, periods_per_year)
		if ann_vol == 0:
			return 0.0
		return (ann_return - risk_free_rate) / ann_vol


	@classmethod
	def analyze_drawdowns(cls, daily_returns):
		drawdowns = []
		peak = 1.0
		value = 1.0
		in_dd = False
		dd_start = None
		dd_trough = None
		dd_trough_value = None
		dd_peak_value = None
		for i, r in enumerate(daily_returns):
			value *= (1 + r)
# 創新高 → 回撤結束
			if value > peak:
				if in_dd:
					drawdowns.append({
						"start": dd_start,
						"trough": dd_trough,
						"end": i,
						"depth": dd_trough_value / dd_peak_value - 1,
						"duration": i - dd_start,
						"recovered": True
					})
					in_dd = False
				peak = value
			else:
# 尚未創新高 → 在回撤中
				if not in_dd:
# 回撤開始
					in_dd = True
					dd_start = i
					dd_trough = i
					dd_trough_value = value
					dd_peak_value = peak
				elif value < dd_trough_value:
# 回撤開始以後，出現更低的低點
					dd_trough = i
					dd_trough_value = value
# 資料結束仍未 recovery
		if in_dd:
			drawdowns.append({
				"start": dd_start,
				"trough": dd_trough,
				"end": None,
				"depth": dd_trough_value / dd_peak_value - 1,
				"duration": len(daily_returns) - dd_start,
				"recovered": False
			})
		return drawdowns


	@classmethod
	def drawdown_summary(cls, drawdowns):
		if not drawdowns:
# 極端的例子 如果從開始就一路上漲 沒有任何回撤
			return {
				"Max Drawdown": 0.0,
				"Max DD Duration": 0,
				"Recovered MaxDD": True,
				"Current DD Duration": 0,
				"Total Drawdowns": 0,
				"Recovered Count": 0,
				"Unrecovered Count": 0
			}
		max_dd_event = min(drawdowns, key=lambda d: d["depth"])
		recovered = [d for d in drawdowns if d["recovered"]]
		unrecovered = [d for d in drawdowns if not d["recovered"]]
		return {
			"Max Drawdown": max_dd_event["depth"],
			"Max DD Duration": max_dd_event["duration"],
			"Recovered MaxDD": max_dd_event["recovered"],
			"Current DD Duration": unrecovered[-1]["duration"] if unrecovered else 0,
			"Total Drawdowns": len(drawdowns),
			"Recovered Count": len(recovered),
			"Unrecovered Count": len(unrecovered)
		}


	def __init__(self, cfg):
		self.xcfg = {
			"source_folderpath": None,
			"source_filename": None,
			"risk_free_rate": 0.0,
			"statistics_date_range_string": None,
			# "date_range_start": None,
			# "date_range_end": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["source_folderpath"] = self.DEFAULT_DATA_FOLDERPATH if self.xcfg["source_folderpath"] is None else self.xcfg["source_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_filename"])
		# print ("__init__: %s" % self.xcfg["source_filepath"])
		self.workbook = None
		self.cur_year = datetime.now().year
		self.worksheet_data = None

		self.filepath_dict = OrderedDict()
		self.filepath_dict["source_filepath"] = self.xcfg["source_filepath"]


	def __enter__(self):
		return self


	def __exit__(self, type, msg, traceback):
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


	def __get_worksheet(self):
		if self.workbook is None:
			# import pdb; pdb.set_trace()
			if not self.__check_file_exist(self.xcfg["source_filepath"]):
				raise RuntimeError("The worksheet[%s] does NOT exist" % self.xcfg["source_filepath"])
			self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
			self.worksheet = self.workbook.sheet_by_index(0)
		return self.worksheet


	def __read_worksheet(self, expected_title_list=None):
# Check if it's required to transform from stock name to stock symbol
		worksheet_data = {}
		# import pdb; pdb.set_trace()			
		data_list = []
# title
		title_list = []
		for column_index in range(0, self.Worksheet.ncols):
			title_value = self.Worksheet.cell_value(0, column_index)
			title_list.append(title_value)
		# print(title_list)
		# import pdb; pdb.set_trace()
		time_index = None
		title_index_list = None
		if expected_title_list is not None:
			# time_index = title_list.index(self.DEFAULT_TIME_FIELD_NAME)
			# title_index_list = [time_index,]
			title_index_list = []
			for expected_title in expected_title_list:
				index = title_list.index(expected_title)
				title_index_list.append(index)
			new_title_list = [title for index, title in enumerate(title_list) if index in title_index_list]
			title_list = new_title_list
		else:
			title_index_list = list(range(0, self.Worksheet.ncols))
		# if time_index is not None:
		# 	raise ValueError("The expected title list must include the time field name: %s" % self.DEFAULT_TIME_FIELD_NAME)
# data
		for row_index in range(1, self.Worksheet.nrows):
			entry_list = []
			can_add = True
			# for column_index in range(0, self.Worksheet.ncols):
			for column_index in title_index_list:
				entry_value = self.Worksheet.cell_value(row_index, column_index)
				if column_index == time_index:
					day_diff = int(entry_value) - self.DEFAULT_DATE_BASE_NUMBER
					entry_date = self.DEFAULT_DATE_BASE + timedelta(days=day_diff)
					if entry_date < date(2010, 1, 1):
						can_add = False
						break
					entry_value = entry_date.strftime("%m/%d/%Y")
					# print("%d: %s" % (row_index + 1, entry_value))
					# import pdb; pdb.set_trace()
				entry_list.append(entry_value)
			# print("%d: %s" % (row_index + 1, entry_list))
			if can_add: data_list.append(entry_list)
		worksheet_data["title"] = title_list
		worksheet_data["data"] = data_list
		return worksheet_data


	def __get_worksheet_data(self, expected_title_list=None):
		need_read = False
		if self.worksheet_data is None:
			need_read = True
		else:
			if expected_title_list is None:
				if self.Worksheet.ncols != len(self.worksheet_data["title"]):
					need_read = True
			else:
				expected_title_list_len = len(expected_title_list)
				if len(self.worksheet_data["title"]) != expected_title_list_len:
					need_read = True
				else:
					for i in range(expected_title_list_len):
						if expected_title_list[i] != self.worksheet_data["title"][i]:
							need_read = True
							break
		if need_read:
			self.worksheet_data = self.__read_worksheet(expected_title_list)
		return self.worksheet_data


	def __date_str2list(self, date_str, skip_year=False):
		# import pdb; pdb.set_trace()
		# print(date_str)
		if date_str.find("/") != -1:
			elem_list = date_str.split("/")
			if len(elem_list) == 2:
				[year, month, day] = [self.cur_year, int(elem_list[0]), int(elem_list[1])]
			elif len(elem_list) == 3:
				[month, day, year] = list(map(int, elem_list))
			else:
				raise ValueError("Incorrect date string format: %s" % date_str)
		elif date_str.find("-") != -1:
			elem_list = date_str.split("-")
			if len(elem_list) == 2:
				[year, month, day] = [self.cur_year, int(elem_list[0]), int(elem_list[1])]
			elif len(elem_list) == 3:
				[year, month, day] = list(map(int, elem_list))
			else:
				raise ValueError("Incorrect date string format: %s" % date_str)
		else:
			raise ValueError("Incorrect date string format: %s" % date_str)
		return [month, day] if skip_year else [year, month, day]


	def __date_str2obj(self, date_str):
		[year, month, day] = self.__date_str2list(date_str)
		date_obj = None
		try:
			date_obj = date(year, month, day)
		except Exception as e:
			print("Unsupport date format[%s] due to: %s" % (date_str, str(e)))
			# import pdb; pdb.set_trace()
			raise e
		return date_obj


	def __date_number2obj(self, date_number):
		day_diff = int(date_number) - self.DEFAULT_DATE_BASE_NUMBER
		date_obj = self.DEFAULT_DATE_BASE + timedelta(days=day_diff)
		return date_obj


	def __date_number2list(self, date_number):
		date_obj = self.__date_number2obj(date_number)
		return [date_obj.year, date_obj.month, date_obj.day]


	def __extract_data(self, date_range_start=None, date_range_end=None):
		# import pdb; pdb.set_trace()
		expedcted_title_list = [self.DEFAULT_TIME_FIELD_NAME, self.DEFAULT_CLOSING_PRICE_FIELD_NAME,]
		worksheet_data = self.__get_worksheet_data(expedcted_title_list)
		time_index = worksheet_data["title"].index(self.DEFAULT_TIME_FIELD_NAME)
		need_check_time_range = (date_range_start is not None) or (date_range_end is not None)
		date_range_start_list = None
		date_range_end_list = None
		if need_check_time_range:
			if date_range_start is not None:
				date_range_start_list = self.__date_str2list(date_range_start)
			else:
				date_range_start_list = self.__date_str2list("01/01/1900")
			if date_range_end is not None:
				date_range_end_list = self.__date_str2list(date_range_end)
			else:
				date_range_end_list = self.__date_str2list("12/31/2100")
			funcptr = lambda x: date_range_start_list <= self.__date_number2list(x[time_index]) <= date_range_end_list
			data_tmp = filter(funcptr, worksheet_data["data"])
			worksheet_data["data"] = list(data_tmp)
		return worksheet_data


	def print_filepath(self):
		print("************** File Path **************")
		for key, value in self.filepath_dict.items():
			print("%s: %s" % (key, value))


	def analyze_performance(self):
		# import pdb; pdb.set_trace()
		date_range_start = None
		date_range_end = None
		if self.xcfg["statistics_date_range_string"] is not None:
			date_range_list = self.xcfg["statistics_date_range_string"].split(":")
			if len(date_range_list) != 2:
				raise ValueError("Incorrect date range format: %s" % self.xcfg["statistics_date_range_string"])
			[date_range_start, date_range_end] = date_range_list
			if date_range_start == "":
				date_range_start = None
			if date_range_end == "":
				date_range_end = None
		worksheet_data = self.__extract_data(date_range_start, date_range_end)
		daily_returns = []
		closing_price_index = worksheet_data["title"].index(self.DEFAULT_CLOSING_PRICE_FIELD_NAME)
		prev_closing_price = float(worksheet_data["data"][0][closing_price_index])
		for row in worksheet_data["data"][1:]:
			closing_price = float(row[closing_price_index])
			# if len(daily_returns) == 0:
			# 	prev_closing_price = closing_price
			# 	continue
			daily_return = (closing_price - prev_closing_price) / prev_closing_price
			daily_returns.append(daily_return)
			prev_closing_price = closing_price
		drawdowns = self.analyze_drawdowns(daily_returns)
		dd_summary = self.drawdown_summary(drawdowns)
		return {
			"Cumulative Return": self.cumulative_return(daily_returns),
			"CAGR": self.cagr(daily_returns),
			"Annualized Volatility": self.annualized_volatility(daily_returns),
			"Sharpe Ratio": self.sharpe_ratio(daily_returns, self.xcfg["risk_free_rate"]),
			"Max Dropdown": dd_summary,
		}


	def show_performance(self):
		PERCENT_KEYS = {"Cumulative Return", "CAGR", "Annualized Volatility", "Max Drawdown",}
		perf_dict = self.analyze_performance()
		print("Performance Analysis:")
		for key, value in perf_dict.items():
			if isinstance(value, float):
				if key in PERCENT_KEYS:
					print("  %s: %.2f%%" % (key, value * 100))
				else:
					print("  %s: %.2f" % (key, value))
			elif isinstance(value, int):
				print("  %s: %s" % (key, value))
			elif isinstance(value, dict):
				if key == "Max Dropdown":
					print("  %s:" % key)
					for dd_key, dd_value in value.items():
						if isinstance(dd_value, float):
							if dd_key in PERCENT_KEYS:
								print("    %s: %.2f%%" % (dd_key, dd_value * 100))
							else:
								print("    %s: %.2f" % (dd_key, dd_value))
						else:
							print("    %s: %s" % (dd_key, dd_value))
				else:
					raise ValueError("Unsupport performance data type: %s" % type(value))
			else:
				raise ValueError("Unsupport performance data type: %s" % type(value))


	@property
	def Worksheet(self):
		return self.__get_worksheet()


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
	'store_true' and 'store_false' - ?些是 'store_const' 分?用作存? True 和 False 值的特殊用例。
	另外，它?的默?值分?? False 和 True。例如:

	>>> parser = argparse.ArgumentParser()
	>>> parser.add_argument('--foo', action='store_true')
	>>> parser.add_argument('--bar', action='store_false')
	>>> parser.add_argument('--baz', action='store_false')
	'''
	parser.add_argument('--source_folderpath', required=False, help='Update database from the XLS files in the designated folder path. Ex: %s' % PerformanceAnalysis.DEFAULT_DATA_FOLDERPATH)
	parser.add_argument('--source_filename', required=False, help='The filename of chip analysis data source')
	parser.add_argument('-s', '--show_performance', required=False, action='store_true', help='Show the result of performace analysis for the specific target and exit.')
	parser.add_argument('--statistics_date_range', required=False, 
		 help='''The statistics data during the date range.
  Date range
    Format: yy1-mm1-dd1:yy2-mm2-dd2   From yy1-mm1-dd1 to yy2-mm2-dd2   Ex: 2014-09-04:2025-10-15
	Format: yy-mm-dd:   From yy-mm-dd to 'the last date of the data'   Ex: 2014-09-04:
	Format: :yy-mm-dd   From 'the first date of the data' to yy-mm-dd   Ex: :2025-09-04
	* Caution: Only take effect when --show_performance is set.''')
	parser.add_argument('--print_filepath', required=False, action='store_true', help='Print the filepaths used in the process and exit.')
	args = parser.parse_args()
	# import pdb; pdb.set_trace()
	cfg = {}
	if args.source_folderpath is not None: cfg['source_folderpath'] = args.source_folderpath
	if args.source_filename is not None: cfg['source_filename'] = args.source_filename
	if args.statistics_date_range is not None: cfg['statistics_date_range_string'] = args.statistics_date_range
	# import pdb; pdb.set_trace()
	with PerformanceAnalysis(cfg) as obj:
		if args.print_filepath:
			obj.print_filepath()
			sys.exit(0)
		if args.show_performance:
			obj.show_performance()
			sys.exit(0)