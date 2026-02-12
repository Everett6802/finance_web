#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import argparse
import errno
# '''
# Question: How to Solve xlrd.biffh.XLRDError: Excel xlsx file; not supported ?
# Answer : The latest version of xlrd(2.01) only supports .xls files. Installing the older version 1.2.0 to open .xlsx files.
# '''
# import xlrd
# import xlsxwriter
from openpyxl import Workbook, load_workbook
import math
import pandas as pd
import csv
import argparse
from datetime import datetime, date, timedelta
import getpass
from collections import OrderedDict
import yfinance as yf

class ReadXLSException(Exception): pass

class DataFetch(object):

	DEFAULT_HOST_DATA_FOLDERPATH =  "C:\\Users\\%s\\project_data\\finance_web" % getpass.getuser()
	DEFAULT_DATA_FOLDERPATH =  os.getenv("DATA_PATH", DEFAULT_HOST_DATA_FOLDERPATH)
	# DEFAULT_SOURCE_FILENAME = "加權指數歷史資料2000-2025.xlsx"
	DEFAULT_TIME_FIELD_NAME = "時間"	
	DEFAULT_CLOSING_PRICE_FIELD_NAME = "收盤價"	
	# DEFAULT_CONFIG_FOLDERPATH =  "C:\\Users\\%s" % os.getlogin()
	# DEFAULT_DATE_BASE_NUMBER = 36526
	# DEFAULT_DATE_BASE = date(2000, 1, 1)
	DEFAULT_YAHOO_DATA_TITLE_MAPPING = [("Date", "時間"), ("Open", "開盤價"), ("High", "最高價"), ("Low", "最低價"), ("Close", "收盤價"), ("Volume", "成交量")]
	DEFAULT_YAHOO_TILE_LIST = [x[0] for x in DEFAULT_YAHOO_DATA_TITLE_MAPPING]
	DEFAULT_DATA_TILE_LIST = [x[1] for x in DEFAULT_YAHOO_DATA_TITLE_MAPPING]
	DEFAULT_YAHOO_DATE_TITLE = "Date"
	DEFAULT_YAHOO_DATE_TITLE_INDEX = DEFAULT_YAHOO_TILE_LIST.index(DEFAULT_YAHOO_DATE_TITLE)
	DEFAULT_YAHOO_OPEN_TITLE = "Open"
	DEFAULT_YAHOO_OPEN_TITLE_INDEX = DEFAULT_YAHOO_TILE_LIST.index(DEFAULT_YAHOO_OPEN_TITLE)
	DEFAULT_YAHOO_HIGH_TITLE = "High"
	DEFAULT_YAHOO_HIGH_TITLE_INDEX = DEFAULT_YAHOO_TILE_LIST.index(DEFAULT_YAHOO_HIGH_TITLE)
	DEFAULT_YAHOO_LOW_TITLE = "Low"
	DEFAULT_YAHOO_LOW_TITLE_INDEX = DEFAULT_YAHOO_TILE_LIST.index(DEFAULT_YAHOO_LOW_TITLE)
	DEFAULT_YAHOO_CLOSE_TITLE = "Close"
	DEFAULT_YAHOO_CLOSE_TITLE_INDEX = DEFAULT_YAHOO_TILE_LIST.index(DEFAULT_YAHOO_CLOSE_TITLE)
	DEFAULT_YAHOO_VOLUME_TITLE = "Volume"
	DEFAULT_YAHOO_VOLUME_TITLE_INDEX = DEFAULT_YAHOO_TILE_LIST.index(DEFAULT_YAHOO_VOLUME_TITLE)
	DEFAULT_YAHOO_DATE_FORMAT = "%Y-%m-%d"
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


	def __init__(self, cfg):
		self.xcfg = {
			"source_folderpath": None,
			"source_filename": None,
			"stock_symbol_string": None,
			"data_date_range_string": None,
			"refresh_data": False,
			# "date_range_start": None,
			# "date_range_end": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)
		self.xcfg["source_folderpath"] = self.DEFAULT_DATA_FOLDERPATH if self.xcfg["source_folderpath"] is None else self.xcfg["source_folderpath"]
		# self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		# self.xcfg["source_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_filename"])
		# print ("__init__: %s" % self.xcfg["source_filepath"])
		self.workbook = None

		self.filepath_dict = OrderedDict()
		self.filepath_dict["source_folderpath"] = self.xcfg["source_folderpath"]


	def __enter__(self):
		return self


	def __exit__(self, type, msg, traceback):
		# if self.workbook is not None:
		# 	self.workbook.close()
		# 	# del self.workbook
		# 	self.workbook = None
		return False


	# def __get_workbook(self):
	# 	if self.workbook is None:
	# 		# if not self.__check_file_exist(self.xcfg["source_filepath"]):
	# 		# 	raise ReadXLSException("The file[%s] does NOT exist" % self.xcfg["source_filepath"])
	# 		self.workbook = load_workbook(self.xcfg["source_filepath"])
	# 	return self.workbook


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


# ========= XLSX 工具 =========
	def __read_xlsx(self, source_filepath):
		# wb = self.__get_workbook()
		wb = load_workbook(source_filepath)
		ws = wb.active
		rows = []
		header = [c.value for c in ws[1]]
		rows.append(header)
		# for row in ws.iter_rows(min_row=2, values_only=True):
		# 	rows.append(dict(zip(header, row)))
		for row in ws.iter_rows(min_row=2, values_only=True):
			rows.append(row)
		return rows


	def __write_xlsx(self, source_filepath, rows):
		ws = None
		start_index = None
		# import pdb; pdb.set_trace()
		if not self.xcfg["refresh_data"] and self.__check_file_exist(source_filepath):
# If file exists, append new data...
			wb = load_workbook(source_filepath)
			ws = wb.active
# openpyxl 的 row / column 是從 1 開始，不是 0
# openpyxl 是 Excel 操作庫，不是資料結構庫 它選擇「跟 Excel 一樣」而不是「跟 Python 一樣」
# Excel 的世界本來就是從 1 開始
			data_last_date_str = ws.cell(row=ws.max_row, column=self.DEFAULT_YAHOO_DATE_TITLE_INDEX + 1).value
			data_last_date = datetime.strptime(data_last_date_str, self.DEFAULT_YAHOO_DATE_FORMAT).date()
			for index, row in enumerate(rows[1:]):
				row_date_str = row[self.DEFAULT_YAHOO_DATE_TITLE_INDEX]
				row_date = datetime.strptime(row_date_str, self.DEFAULT_YAHOO_DATE_FORMAT).date()
				if row_date > data_last_date:
# row 0 是 header，所以 index 從 0 開始對應到 rows[1:] 的第一行資料
					start_index = index + 1 # 因為 rows[1:] 的 index 是從 0 開始，所以要加 1 才是正確的行數
					break
			# if start_index is None: # 已存在，不寫入
			# 	wb.close()
			# 	return
		else:
# If file does NOT exist, create new file...
			wb = Workbook()
			ws = wb.active
			# headers = rows[0].keys()
			# ws.append(list(headers))
			ws.append(self.DEFAULT_DATA_TILE_LIST)
			start_index = 1
		# import pdb; pdb.set_trace()
		if start_index != None:
			for r in rows[start_index:]:
				ws.append(r)
			data_len = len(rows[start_index:])
			print(f"{data_len} is updated in {source_filepath}")
			wb.save(source_filepath)
		else:
			print(f"No data is updated in {source_filepath}")
		wb.close()


	def __fetch_and_cache(self, stock_symbol, date_range_start_str=None, date_range_end_str=None):
		"""
		取得歷史資料，如果本地已存在 CSV，則只抓最新日期後的資料。
		"""
		# csv_file = os.path.join(self.xcfg["source_filepath"], f"{stock_symbol}.csv")
# 檢查是否已存在本地資料
		# import pdb; pdb.set_trace()
		source_filepath = os.path.join(self.xcfg["source_folderpath"], f"{stock_symbol}.xlsx")
		file_exist = self.__check_file_exist(source_filepath)
		if file_exist:
			if self.__is_excel_locked(source_filepath):
				print(f"WARNING: The file {source_filepath} is locked by other process, so skip fetching data for {stock_symbol}...")
				return False
		fetch_start = fetch_end = None
		if not self.xcfg["refresh_data"] and file_exist:
			rows = self.__read_xlsx(source_filepath)
			last_date_str = rows[-1][self.DEFAULT_YAHOO_DATE_TITLE_INDEX]
			last_date = datetime.strptime(last_date_str, self.DEFAULT_YAHOO_DATE_FORMAT).date()
			fetch_start = last_date + timedelta(days=1)
			if date_range_start_str is not None:
				date_range_start = datetime.strptime(date_range_start_str, self.DEFAULT_YAHOO_DATE_FORMAT).date()
				if date_range_start > fetch_start:
					print(f"WARNING: The start date {date_range_start_str} is later than {last_date_str} in local data, so no data will be fetched.")
					return False
		else:
			if date_range_start_str is None:
				fetch_start = self.DEFAULT_DATE_BASE
			else:
				fetch_start = datetime.strptime(date_range_start_str, self.DEFAULT_YAHOO_DATE_FORMAT).date()
		if date_range_end_str is None:
			fetch_end = datetime.today().date()
		else:
			fetch_end = datetime.strptime(date_range_end_str, self.DEFAULT_YAHOO_DATE_FORMAT).date()
			if fetch_end > datetime.today().date():
				print(f"WARNING: The end date {date_range_start_str} shuld NOT be later than today")
				return False
# 如果已經最新，直接返回
		if fetch_start > fetch_end:
			print(f"WARNING: Incorrect time range %s - %s" % (fetch_start.strftime(self.DEFAULT_YAHOO_DATE_FORMAT), fetch_end.strftime(self.DEFAULT_YAHOO_DATE_FORMAT)))
			return False
# 抓取歷史資料
		hist = yf.download(stock_symbol, start=fetch_start.strftime(self.DEFAULT_YAHOO_DATE_FORMAT), end=fetch_end.strftime(self.DEFAULT_YAHOO_DATE_FORMAT), group_by='column')
		hist = hist.reset_index()
# 把 MultiIndex 欄位轉成單層欄位
		if isinstance(hist.columns, pd.MultiIndex):
			hist.columns = hist.columns.get_level_values(0)
# 轉成 CSV 格式
		row_data_list = []
		row_data_list.append(self.DEFAULT_DATA_TILE_LIST)
		# import pdb; pdb.set_trace()
		for _, r in hist.iterrows():
			one_row_data = []
			for data_title in self.DEFAULT_YAHOO_TILE_LIST:
				if data_title == self.DEFAULT_YAHOO_DATE_TITLE:
					date_str = r[data_title].strftime(self.DEFAULT_YAHOO_DATE_FORMAT)
					one_row_data.append(date_str)
				else:
					one_row_data.append(r[data_title])
# Check if fake row
			if one_row_data[self.DEFAULT_YAHOO_VOLUME_TITLE_INDEX] == 0:
				[o, h, l, c] = [one_row_data[i] for i in [self.DEFAULT_YAHOO_OPEN_TITLE_INDEX, self.DEFAULT_YAHOO_HIGH_TITLE_INDEX, self.DEFAULT_YAHOO_LOW_TITLE_INDEX, self.DEFAULT_YAHOO_CLOSE_TITLE_INDEX]]
				if not any(math.isnan(x) for x in [o,h,l,c]):
					if o == h == l == c:
						continue
			row_data_list.append(one_row_data)
# 寫入 XLSX
		# import pdb; pdb.set_trace()
		self.__write_xlsx(source_filepath, row_data_list)
		return True


	def __is_excel_locked(self, filepath):
# Excel lock file
		lock_path = os.path.join(os.path.dirname(filepath), "~$" + os.path.basename(filepath))
		if os.path.exists(lock_path):
			return True
# OS lock
		try:
			with open(filepath, "a"):
				return False
		except PermissionError:
			return True


	def fetch_data(self):
		# import pdb; pdb.set_trace()
		if self.xcfg["stock_symbol_string"] is None:
			print("Warning: No stock to fetch...")
			return
		stock_symbol_list = self.xcfg["stock_symbol_string"].split(",")
		date_range_start_str = date_range_end_str = None
		if self.xcfg["data_date_range_string"] is not None:
			date_range_elems = self.xcfg["data_date_range_string"].split(":")
			if len(date_range_elems) == 2:
				if date_range_elems[0] != "":
					date_range_start_str = date_range_elems[0]
				if date_range_elems[1] != "":
					date_range_end_str = date_range_elems[1]
			else:
				print("Error: Incorrect date range format[%s]" % self.xcfg["data_date_range_string"])
				return 
		for stock_symbol in stock_symbol_list:
			fetch_success = self.__fetch_and_cache(stock_symbol, date_range_start_str, date_range_end_str)
			if not fetch_success:
				print(f"Fails to fetch {stock_symbol} data...")


	def __inspect_data(self, stock_symbol):
		# import pdb; pdb.set_trace()
		source_filepath = os.path.join(self.xcfg["source_folderpath"], f"{stock_symbol}.xlsx")
		file_exist = self.__check_file_exist(source_filepath)
		data_info = None
		if file_exist:
			if self.__is_excel_locked(source_filepath):
				# print(f"WARNING: The file {source_filepath} is locked by other process, fails to show data info for {stock_symbol}...")
				raise ReadXLSException(f"The file {source_filepath} is locked by other process, fails to inspect data info for {stock_symbol}...")
			rows = self.__read_xlsx(source_filepath)
			data_info = {
				"first_date_str": rows[1][self.DEFAULT_YAHOO_DATE_TITLE_INDEX], 
				"last_date_str": rows[-1][self.DEFAULT_YAHOO_DATE_TITLE_INDEX], 
				"data_cnt": len(rows) - 1,
			}
		return data_info


	def show_data_info(self):
		# import pdb; pdb.set_trace()
		if self.xcfg["stock_symbol_string"] is None:
			print("Warning: No stock to fetch...")
			return
		stock_symbol_list = self.xcfg["stock_symbol_string"].split(",")
		for stock_symbol in stock_symbol_list:
			try:
				data_info = self.__inspect_data(stock_symbol)
				if data_info is not None:
					first_date_str = data_info["first_date_str"]
					last_date_str = data_info["last_date_str"]
					data_cnt = data_info["data_cnt"]
					print(f"{stock_symbol}: {data_cnt} data from {first_date_str} to {last_date_str}")
				else:
					print(f"{stock_symbol}: No data exists...")
			except ReadXLSException as e:
				print(str(e))


	def print_filepath(self):
		print("************** File Path **************")
		for key, value in self.filepath_dict.items():
			print("%s: %s" % (key, value))


if __name__ == "__main__":
# argparse 預設會把 help 文字裡的換行與多重空白「壓縮」成一行，所以你在字串裡寫的 \n 不一定會照原樣顯示。 => 建立 parser 時加上 formatter_class=argparse.RawTextHelpFormatter
	parser = argparse.ArgumentParser(description='Print help', formatter_class=argparse.RawTextHelpFormatter)
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
	parser.add_argument('--source_folderpath', required=False, help='Fetch data into the XLS files in the designated folder path. Ex: %s' % DataFetch.DEFAULT_DATA_FOLDERPATH)
	# parser.add_argument('--source_filename', required=False, help='The filename of chip analysis data source')
	parser.add_argument('-f', '--fetch_data', required=False, action='store_true', help='Fetch the data of the specific target and exit.')
	parser.add_argument('--stock_symbol_list', required=False, help='The stock symbol list. Mutiple stock symbols are split by comma. Ex: 00850.TW,00881.TW,00692.TW,MSFT,GOOG')
	parser.add_argument('--data_date_range', required=False, 
		 help='''The data during the date range.
  Date range:
    Format: yy1-mm1-dd1:yy2-mm2-dd2   From yy1-mm1-dd1 to yy2-mm2-dd2   Ex: 2014-09-04:2025-10-15
    Format: yy-mm-dd:   From yy-mm-dd to 'the last date of the data'   Ex: 2014-09-04:
    Format: :yy-mm-dd   From 'the first date of the data' to yy-mm-dd   Ex: :2025-09-04
    * Caution: Only take effect when --fetch_data is set.''')
	parser.add_argument('--refresh_data', required=False, action='store_true', help='Ignore the existing the XLSX file so that the data will be fetched from the earliest date to today. Only take effect when --fetch_data is set.')
	parser.add_argument('--show_data_info', required=False, action='store_true', help='Show data info and exit.')
	parser.add_argument('--print_filepath', required=False, action='store_true', help='Print the filepaths used in the process and exit.')
	args = parser.parse_args()
	# import pdb; pdb.set_trace()
	cfg = {}
	if args.source_folderpath is not None: cfg['source_folderpath'] = args.source_folderpath
	if args.stock_symbol_list is not None: cfg['stock_symbol_string'] = args.stock_symbol_list
	if args.data_date_range is not None: cfg['data_date_range_string'] = args.data_date_range
	if args.refresh_data: cfg['refresh_data'] = True
	# import pdb; pdb.set_trace()
	with DataFetch(cfg) as obj:
		if args.show_data_info:
			obj.show_data_info()
			sys.exit(0)
		if args.print_filepath:
			obj.print_filepath()
			sys.exit(0)
		if args.fetch_data:
			obj.fetch_data()
			sys.exit(0)