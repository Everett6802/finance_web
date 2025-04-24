#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import xlrd
import json
# import xlsxwriter
import argparse
import errno
import time
# from datetime import datetime
import datetime
from collections import OrderedDict


class ScrapyError(Exception): pass

class TakeProfitTracker(object):

	DEFAULT_DATA_FOLDERPATH =  "C:\\停利追蹤"
	DEFAULT_SOURCE_FILENAME = "take_profile_tracker"
	DEFAULT_SOURCE_FULL_FILENAME = "%s.xlsx" % DEFAULT_SOURCE_FILENAME
	DEFAULT_RECORD_FILENAME = "take_profile_tracker_record"
	DEFAULT_RECORD_FULL_FILENAME = "%s.txt" % DEFAULT_RECORD_FILENAME
	DEFAULT_CUSTOMIZED_CONFIG_FILENAME = "take_profile_tracker_customized_config"
	DEFAULT_CUSTOMIZED_CONFIG_FULL_FILENAME = "%s.json" % DEFAULT_CUSTOMIZED_CONFIG_FILENAME
	DEFAULT_HTML_RESULT_FILENAME = "take_profile_tracker_result"
	DEFAULT_HTML_RESULT_FULL_FILENAME = "%s.html" % DEFAULT_HTML_RESULT_FILENAME
	DEFAULT_STOCK_SYMBOL_LOOKUP_FILENAME = "股號查詢"
	DEFAULT_STOCK_SYMBOL_LOOKUP_FULL_FILENAME = "%s.xlsx" % DEFAULT_STOCK_SYMBOL_LOOKUP_FILENAME
	DEFAULT_TRAILING_STOP_RATIO = 0.7
	DEFAULT_TRIGGER_TRAILING_STOP_PROFIT_RATIO = 0.15
# 商品,平圴成本,股數,最大獲利,停利價格,啟動停利
	DEFAULT_RECORD_FIELD_METADATA = [["商品", str], ["平圴成本", float], ["股數", int], ["最大獲利", int], ["停利價格", float], ["啟動停利", str]]  # , ['獲利%', float]
	DEFAULT_RECORD_FIELD_NAME = [metadata[0] for metadata in DEFAULT_RECORD_FIELD_METADATA]
	DEFAULT_RECORD_FIELD_TYPE = [metadata[1] for metadata in DEFAULT_RECORD_FIELD_METADATA]
	DEFAULT_RECORD_FIELD_BRAND_NEW_NAME_LEN = 3
	DEFAULT_RECORD_FIELD_BRAND_NEW_NAME = DEFAULT_RECORD_FIELD_NAME[0:DEFAULT_RECORD_FIELD_BRAND_NEW_NAME_LEN]
	DEFAULT_RECORD_FIELD_EXISTING_DATA_NAME_LEN = 4
	DEFAULT_RECORD_FIELD_EXISTING_DATA_NAME = DEFAULT_RECORD_FIELD_NAME[0:DEFAULT_RECORD_FIELD_EXISTING_DATA_NAME_LEN]
	DEFAULT_RECORD_FIELD_METADATA_LEN = len(DEFAULT_RECORD_FIELD_METADATA)
	DEFAULT_SHOW_TRACK_FIELD_NAME = ['商品', '漲跌', '漲幅%', "股數", '獲利%', "平圴成本", "最大獲利", "停利價格", '成交', '價差', '價差%']
	DEFAULT_SHOW_TRACK_FIELD_NAME_LEN = len(DEFAULT_SHOW_TRACK_FIELD_NAME)
	YAHOO_STOCK_URL_FORMAT = "https://tw.stock.yahoo.com/quote/%s.TW"
	DEFAULT_MONITOR_TIME_INTERVAL = 300
	DEFAULT_CAN_SCRAPE_TIME_RANGE_START = datetime.time(8, 59, 0)
	DEFAULT_CAN_SCRAPE_TIME_RANGE_END = datetime.time(13, 36, 0)


	def __init__(self, cfg):
		self.xcfg = {
			"data_folderpath": None,
			"source_filename": self.DEFAULT_SOURCE_FULL_FILENAME,
			"record_filename": self.DEFAULT_RECORD_FULL_FILENAME,
			"customized_config_filename": self.DEFAULT_CUSTOMIZED_CONFIG_FULL_FILENAME,
			"html_result_filename": self.DEFAULT_HTML_RESULT_FULL_FILENAME,
			"stock_symbol_lookup_filename": self.DEFAULT_STOCK_SYMBOL_LOOKUP_FULL_FILENAME,
			"trailing_stop_ratio": self.DEFAULT_TRAILING_STOP_RATIO,
			"trigger_trailing_stop_profit_ratio": self.DEFAULT_TRIGGER_TRAILING_STOP_PROFIT_RATIO,
			"read_from_scrapy": False,
			"force_update_record": False,
			"monitor_mode": False,
			"monitor_time_interval": self.DEFAULT_MONITOR_TIME_INTERVAL,
			"show_result": False,
			"output_html_result": False,
			"show_scrapy_progress": False,
		}
		self.xcfg.update(cfg)
		self.xcfg["data_folderpath"] = self.DEFAULT_DATA_FOLDERPATH if self.xcfg["data_folderpath"] is None else self.xcfg["data_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FULL_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["source_filename"])
		self.xcfg["record_filename"] = self.DEFAULT_RECORD_FULL_FILENAME if self.xcfg["record_filename"] is None else self.xcfg["record_filename"]
		self.xcfg["record_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["record_filename"])
		self.xcfg["customized_config_filename"] = self.DEFAULT_CUSTOMIZED_CONFIG_FULL_FILENAME if self.xcfg["customized_config_filename"] is None else self.xcfg["customized_config_filename"]
		self.xcfg["customized_config_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["customized_config_filename"])
		self.xcfg["html_result_filename"] = self.DEFAULT_HTML_RESULT_FULL_FILENAME if self.xcfg["html_result_filename"] is None else self.xcfg["html_result_filename"]
		self.xcfg["html_result_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["html_result_filename"])
		self.xcfg["stock_symbol_lookup_filename"] = self.DEFAULT_STOCK_SYMBOL_LOOKUP_FULL_FILENAME if self.xcfg["stock_symbol_lookup_filename"] is None else self.xcfg["stock_symbol_lookup_filename"]
		self.xcfg["stock_symbol_lookup_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["stock_symbol_lookup_filename"])

		self.workbook = None
		self.worksheet = None
		self.can_lookup_stock_symbol = False
		self.stock_symbol_lookup_dict = None  # 股名 -> 股號
		self.stock_symbol_reverse_lookup_dict = None  # 股號 -> 股名
		self.__read_stock_symbol_mapping_table()
		self.can_scrape = self.__can_scrape()
		self.requests_module = None
		self.beautifulsoup_class = None
		self.stock_data_dict = None
		self.customized_config_dict = None
		self.first_track = True

		self.filepath_dict = OrderedDict()
		self.filepath_dict["source"] = self.xcfg["source_filepath"]
		self.filepath_dict["record"] = self.xcfg["record_filepath"]
		self.filepath_dict["stock_symbol_lookup"] = self.xcfg["stock_symbol_lookup_filepath"]


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
	def __get_line_list_from_file(cls, filepath, startswith=None):
		# import pdb; pdb.set_trace()
		if not cls.__check_file_exist(filepath):
			raise RuntimeError("The file[%s] does NOT exist" % filepath)
		line_list = []
		with open(filepath, 'r') as fp:
			for line in fp:
				if startswith is None:
					if line.startswith("#"): continue
				else:
					if not line.startswith("#"): continue
				line = line.strip("\n")
				if len(line) == 0: continue
				line_list.append(line)
		return line_list


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
	def __get_cur_time(cls, time_only=False):
		cur_time = datetime.datetime.now()
		if time_only:
			cur_time = datetime.time(cur_time.hour, cur_time.minute, cur_time.second)
		return cur_time  # datetime.datetime.now()


	@classmethod
	def __get_cur_timestr(cls):
		return cls.__get_cur_time().strftime('%Y-%m-%d %H:%M:%S')


	@classmethod
	def __check_time_in_range(cls, time_range_start, time_range_end, time_check):
		if time_range_start <= time_range_end:
			return time_range_start <= time_check <= time_range_end
		else:
			return time_range_start <= time_check or time_check <= time_range_end


	@classmethod
	def __float(cls, float_value):
		assert type(float_value) in [int, float], "Incorrect value type: %s" % type(float_value)
		return float("%.2f" % float_value)


	@classmethod
	def __is_trailing_stop_triggered(cls, value):
		mobj = None
		try: 
			mobj = re.match("O", value, re.I)
		except TypeError:
			return False
		return True if (mobj is not None) else False


	@classmethod
	def __get_file_modification_date(cls, filepath):
		if not cls.__check_file_exist(filepath):
			raise ValueError("The file[%s] does NOT exist" % filepath)
		modification_time = os.path.getmtime(filepath)
		# print(modification_time)
		modification_date = datetime.datetime.fromtimestamp(modification_time)
		return modification_date


	@classmethod
	def __get_file_modification_date_str(cls, filepath):
		import pdb; pdb.set_trace()
		file_modification_date = cls.__get_file_modification_date(filepath)
		return file_modification_date.strftime("%Y/%m/%d %H:%M:%S")


	def __enter__(self):
# # Open the workbook
# 		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
# 		self.worksheet = self.workbook.sheet_by_index(0)
		return self


	def __exit__(self, type, msg, traceback):
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.worksheet = None
			self.workbook = None
		return False


	def __get_worksheet(self):
		if self.workbook is None:
			if not self.__check_file_exist(self.xcfg["source_filepath"]):
				raise RuntimeError("The worksheet[%s] does NOT exist" % self.xcfg["source_filepath"])
			self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
			self.worksheet = self.workbook.sheet_by_index(0)
		return self.worksheet


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


	def __read_worksheet(self, stock_id_list=None):
# Check if it's required to transform from stock name to stock symbol
		worksheet_data = {}
		# import pdb; pdb.set_trace()
		title_list = []
# title
		for column_index in range(1, self.Worksheet.ncols):
			title_value = self.Worksheet.cell_value(0, column_index)
			title_list.append(title_value)
		# print(title_list)
		# import pdb; pdb.set_trace()
		need_lookup_stock_symbol = False
		if re.match(r"[\d]{4,}", self.Worksheet.cell_value(1, 0)) is None:
			if not self.can_lookup_stock_symbol:
				raise RuntimeError("No stock symbol lookup table !!!")
			else:
				need_lookup_stock_symbol = True
# data
		for row_index in range(1, self.Worksheet.nrows):
			data_key = self.Worksheet.cell_value(row_index, 0)
			if need_lookup_stock_symbol:
				data_key = self.stock_symbol_lookup_dict[data_key]
			if (stock_id_list is not None) and (data_key not in stock_id_list):
				continue
			data_list = []
			for column_index in range(1, self.Worksheet.ncols):
				data_value = self.Worksheet.cell_value(row_index, column_index)
				data_list.append(data_value)
			# print("%s: %s" % (data_key, data_list))
			data_dict = dict(zip(title_list, data_list))
			worksheet_data[data_key] = data_dict
		return worksheet_data


	def __read_record(self):
# 商品,平圴成本,股數,最大獲利,停利價格
		line_list = self.__get_line_list_from_file(self.xcfg["record_filepath"])
		record_data_dict = {}
		# import pdb; pdb.set_trace()
		title_list = line_list[0].split(",")
		title_list_len = len(title_list)
		for line in line_list[1:]:
			line_data_list = line.split(",")
			line_data_list_len = len(line_data_list)
			if line_data_list_len < title_list_len:
				len_diff = title_list_len - line_data_list_len
				line_data_list.extend([None,] * len_diff)
			for index, line_data in enumerate(line_data_list):
				if line_data is not None:
					data_type = self.DEFAULT_RECORD_FIELD_TYPE[index]
					line_data_list[index] = data_type(line_data)
			new_entry = dict(zip(title_list[1:], line_data_list[1:]))
			if line_data_list[0] in record_data_dict:
# Duplicate entry, need to merge the data
				# import pdb; pdb.set_trace()
				old_entry = record_data_dict[line_data_list[0]]
				if new_entry["停利價格"] is not None or new_entry["啟動停利"] is not None:
					raise ValueError("Incorrect value in new entry: %s" % new_entry)
				if old_entry["停利價格"] is not None or old_entry["啟動停利"] is not None:
					raise ValueError("Incorrect value in old entry: %s" % old_entry)
				max_profit_list = []
				if old_entry["最大獲利"] is not None:
					max_profit_list.append(old_entry["最大獲利"])
				if new_entry["最大獲利"] is not None:
					max_profit_list.append(new_entry["最大獲利"])
				if len(max_profit_list) != 0:
					old_entry["最大獲利"] = sum(max_profit_list)
				old_entry["股數"] = old_entry["股數"] + new_entry["股數"]
				old_entry["平圴成本"] = self.__float((old_entry["股數"] * old_entry["平圴成本"] + new_entry["股數"] * new_entry["平圴成本"]) / (old_entry["股數"] + new_entry["股數"]))
			else:
				record_data_dict[line_data_list[0]] = new_entry
		# import pdb; pdb.set_trace()
		return record_data_dict


	def __write_record(self, record_data_dict):
# 商品,平圴成本,股數,最大獲利,停利價格,啟動停利
		# import pdb; pdb.set_trace()
		skipped_line_list = self.__get_line_list_from_file(self.xcfg["record_filepath"], startswith="#")
		with open(self.xcfg['record_filepath'], 'w') as fp:
			line = ",".join(self.DEFAULT_RECORD_FIELD_NAME)
			fp.write("%s\n" % line)
			for stock_symbol in record_data_dict.keys():

				line_data_list = [stock_symbol,]
				line_data_list.append(record_data_dict[stock_symbol]["平圴成本"])
				line_data_list.append(record_data_dict[stock_symbol]["股數"])
				# line_data_list.append(record_data_dict[stock_symbol]["獲利%"])
				line_data_list.append(record_data_dict[stock_symbol]["最大獲利"])
				line_data_list.append(record_data_dict[stock_symbol]["停利價格"])
				line_data_list.append(record_data_dict[stock_symbol]["啟動停利"])
				line_data_list = map(str, line_data_list)
				line = ",".join(line_data_list)
				fp.write("%s\n" % line)
		# import pdb; pdb.set_trace()
		if len(skipped_line_list) != 0:
			with open(self.xcfg['record_filepath'], 'a+') as fp:
				for skipped_line in skipped_line_list:
					fp.write("%s\n" % skipped_line)


	def __read_customized_config(self):
		self.customized_config_dict = {}
		if self.__check_file_exist(self.xcfg["customized_config_filepath"]):
			with open(self.xcfg["customized_config_filepath"], "r") as f:
				self.customized_config_dict = json.load(f)
		return self.customized_config_dict


	def __read_stock_symbol_mapping_table(self):
# 商品,商品,股本
		if self.can_lookup_stock_symbol: return
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(self.xcfg["stock_symbol_lookup_filepath"]):
			print("WARNING: The stock symbol mapping file[%s] does NOT exist" % self.xcfg["stock_symbol_lookup_filepath"])
			return
		stock_symbol_lookup_workbook = xlrd.open_workbook(self.xcfg["stock_symbol_lookup_filepath"])
		stock_symbol_lookup_worksheet = stock_symbol_lookup_workbook.sheet_by_index(0)
# data
		self.stock_symbol_lookup_dict = {}
		self.stock_symbol_reverse_lookup_dict = {}
		for row_index in range(1, stock_symbol_lookup_worksheet.nrows):
			stock_symbol = stock_symbol_lookup_worksheet.cell_value(row_index, 0)
			stock_name = stock_symbol_lookup_worksheet.cell_value(row_index, 1)
			self.stock_symbol_lookup_dict[stock_name] = stock_symbol
			self.stock_symbol_reverse_lookup_dict[stock_symbol] = stock_name
		self.can_lookup_stock_symbol = True
		if stock_symbol_lookup_workbook is not None:
			stock_symbol_lookup_workbook.release_resources()
			del stock_symbol_lookup_workbook
			stock_symbol_lookup_workbook = None


	def __scrape_stock_price(self, stock_symbol):
		# print("Scrape %s: %s" % (stock_symbol, self.__get_cur_timestr()))
		# start_datetime = self.__get_cur_time()
		url = self.YAHOO_STOCK_URL_FORMAT % stock_symbol
		resp = self.__get_requests_module().get(url)
		if re.search(stock_symbol, resp.text) is None:
			raise ValueError("The stock[%s] does NOT exist" % stock_symbol)
		# print(resp.text)
		# import pdb; pdb.set_trace()
		beautifulsoup_class = self.__get_beautifulsoup_class()
		soup = beautifulsoup_class(resp.text, "html.parser")
		div = soup.find("div", {"id": "main-2-QuoteOverview-Proxy"})
		ul = div.find("ul", {"class": "D(f) Fld(c) Flw(w) H(192px) Mx(-16px)"})
		lis = ul.find_all("li")
		# for index, li in enumerate(lis):
		# 	print("================== %d ==================" % index)
		# 	spans = li.find_all("span")
		# 	print(spans[0].text + ": " + spans[1].text)
		def get_value(list_index):
			spans = lis[list_index].find_all("span")
			# print(spans[1].text)
			return spans[1].text
		single_stock_data_dict = {}
		single_stock_data_dict["成交"] = float(get_value(0))
		single_stock_data_dict["漲跌"] = float(get_value(8))
		single_stock_data_dict["漲幅%"] = float(get_value(7).strip("%"))
		is_negative = True if (float(single_stock_data_dict["成交"]) < float(get_value(6))) else False
		if is_negative:
			# single_stock_data_dict["漲跌"] = "-" + single_stock_data_dict["漲跌"]
			# single_stock_data_dict["漲幅%"] = "-" + single_stock_data_dict["漲幅%"]
			single_stock_data_dict["漲跌"] = -1 * single_stock_data_dict["漲跌"]
			single_stock_data_dict["漲幅%"] = -1 * single_stock_data_dict["漲幅%"]
		# print("Scrape %s: %s ... Done" % (stock_symbol, self.__get_cur_timestr()))
		# end_datetime = self.__get_cur_time()
		# print("Scrape %s, Time elaped: %s" % (stock_symbol, (end_datetime - start_datetime)))
		return single_stock_data_dict


	def __read_scrapy(self, stock_id_list):
		stock_data_dict = {}
		# import pdb; pdb.set_trace()
		stock_id_list_len = len(stock_id_list)
		for index, stock_id in enumerate(stock_id_list):
			if self.xcfg["show_scrapy_progress"]:
				start_datetime = self.__get_cur_time()
			stock_data_dict[stock_id] = self.__scrape_stock_price(stock_id)
			if self.xcfg["show_scrapy_progress"]:
				end_datetime = self.__get_cur_time()
				diff_in_sec = (end_datetime - start_datetime).total_seconds()
				progress_percent = (index + 1) * 100.0 / stock_id_list_len
				# print("Scrape %s Done...... %.0f, Time elaped: %.2f(s)" % (stock_id, progress_percent, diff_in_sec))
				print("Scrape {0} Done...... {1:.2f}%, Time elaped: {2:.2f}(s)".format(stock_id, progress_percent, diff_in_sec))
		return stock_data_dict


	def __read_data(self):
		record_data_dict = self.__read_record()
		if self.xcfg["read_from_scrapy"]:
			try:
				self.stock_data_dict = self.__read_scrapy(stock_id_list=record_data_dict.keys())
			except Exception as e:
				print("Scrapy Error: %s" % str(e))
				raise ScrapyError()
		else:
			self.stock_data_dict = self.__read_worksheet(stock_id_list=record_data_dict.keys())
		for key, value in self.stock_data_dict.items():
			value.update(record_data_dict[key])


	def refresh_data(self):
		self.stock_data_dict = None


	def __calculate_trailing_stop_price(self, stock_id, stock_value):
		trailing_stop_ratio = self.xcfg["trailing_stop_ratio"]
		if (stock_id in self.customized_config_dict) and ("trailing_stop_ratio" in self.customized_config_dict[stock_id]):
			trailing_stop_ratio = self.customized_config_dict[stock_id]["trailing_stop_ratio"]
		tmp = stock_value["最大獲利"] * trailing_stop_ratio / stock_value["股數"] + stock_value["平圴成本"]
		return self.__float(tmp)


	def track(self):
		# import pdb; pdb.set_trace()
# update() doesn't return any value (returns None).
		# stock_data_dict = [(key, value, record_data_dict[key], value.update(record_data_dict[key])) for key, value in stock_data_dict.items()]
		# stock_data_dict.update(record_data_dict)
		if self.customized_config_dict is None: self.__read_customized_config()
		if self.stock_data_dict is None: self.__read_data()
		need_update_record = False
		# import pdb; pdb.set_trace()
		take_profit_list = []
		loss_list = []
# 商品,平圴成本,股數,最大獲利,停利價格,啟動停利
		for key, value in self.stock_data_dict.items():
			# value.update(record_data_dict[key])
			if value["成交"] - value["平圴成本"] > 0:
				# import pdb; pdb.set_trace()
				profit = int((value["成交"] - value["平圴成本"]) * value["股數"])
				profit_ratio = profit / (value["平圴成本"] * value["股數"])
				trigger_trailing_stop_profit_ratio = self.xcfg["trigger_trailing_stop_profit_ratio"]
				if (key in self.customized_config_dict) and ("trigger_trailing_stop_profit_ratio" in self.customized_config_dict[key]):
					trigger_trailing_stop_profit_ratio = self.customized_config_dict[key]["trigger_trailing_stop_profit_ratio"]
				should_trigger = profit_ratio > trigger_trailing_stop_profit_ratio
				value["獲利%"] = self.__float(profit_ratio * 100)

				data_changed = False
				if value["停利價格"] is None:
					if value["最大獲利"] is None:
# Brand New
						value["最大獲利"] = profit
					else:
# Exising Data
						if profit > value["最大獲利"]:
							value["最大獲利"] = profit
					if value["啟動停利"] is not None:
						raise ValueError("啟動停利 is NOT None")

					value["啟動停利"] = "O" if should_trigger else "X"
					value["停利價格"] = self.__calculate_trailing_stop_price(key, value)
					data_changed = True
				else:
					if value["最大獲利"] is None:
						raise ValueError("最大獲利 is None, but 停利價格 is NOT None")
					else:
						if profit > value["最大獲利"]:
							value["最大獲利"] = profit
							value["停利價格"] = self.__calculate_trailing_stop_price(key, value)
							data_changed = True
						if not self.__is_trailing_stop_triggered(value["啟動停利"]) and should_trigger:
							value["啟動停利"] = "O"
							data_changed = True
				need_update_record = need_update_record or data_changed
				if self.__is_trailing_stop_triggered(value["啟動停利"]) and value["成交"] < value["停利價格"]:
					# print("停利: %s" % key)
					take_profit_list.append(key)
			else:
				value["獲利%"] = 0.00
				if value["最大獲利"] is None:
# Initial update
					value["最大獲利"] = 0
					value["停利價格"] = 0.00
					value["啟動停利"] = "X"
					need_update_record = True
				else:
					# print("虧損: %s" % key)
					loss_list.append(key)
		if self.xcfg["force_update_record"]: need_update_record = True
		if need_update_record:
			self.__write_record(self.stock_data_dict)
		# print(stock_data_dict)
		cur_time_string = obj.CurTimeString
		print("Data Time: %s" % cur_time_string)
		if self.xcfg["show_result"]:
			self.__show_result()
		if len(take_profit_list) != 0 or len(loss_list) != 0:
			print("\n************************************************")
			if len(take_profit_list) != 0:
				print("停利: %s" % " ".join(take_profit_list))
			if len(loss_list) != 0:
				print("虧損: %s" % " ".join(loss_list))
			print("************************************************\n")
		if self.xcfg["output_html_result"]:
			self.__output_html_result(cur_time_string, take_profit_list, loss_list)


	def __show_result(self):
		if self.stock_data_dict is None: self.__read_data()
# ['商品', '漲跌', '漲幅%', "股數", '獲利%', "平圴成本", "最大獲利", "停利價格", '成交', '價差', '價差%']
		# print("  ".join(self.DEFAULT_SHOW_TRACK_FIELD_NAME))
		print("  ".join(map(lambda x: "%4s" % x, self.DEFAULT_SHOW_TRACK_FIELD_NAME)))
		# import pdb; pdb.set_trace()
		for key, value in self.stock_data_dict.items():
			# value.update(record_data_dict[key])
			data_list = [key,]
			# import pdb; pdb.set_trace()
			for field_name in self.DEFAULT_SHOW_TRACK_FIELD_NAME[1:9]:
				data_list.append(value[field_name])
			# import pdb; pdb.set_trace()
			diff_value = 0.00
			diff_value_percentage = 0.00
			if value["成交"] - value["平圴成本"] > 0:
				diff_value = self.__float(value['成交'] - value['停利價格'])
				diff_value_percentage = self.__float(diff_value / value['停利價格'] * 100.0)
			data_list.extend([diff_value, diff_value_percentage,])
			# print("  ".join(map(str, data_list)))
			str_tmp = "  ".join(map(lambda x: "%8s" % str(x), data_list))
			marker = "* " if self.__is_trailing_stop_triggered(value['啟動停利']) else "  "
			print(marker + str_tmp)


	def __output_html_result(self, data_time, take_profit_list,loss_list):
		with open(self.xcfg['html_result_filepath'], 'w') as fp:
			def add_table_row(fp, line_list):
				fp.write('<tr><td>')
				# print('    </td><td>     '.join(line_list))
				fp.write('</td><td>'.join(line_list))
				fp.write('</td></tr>')
			# fp.write('<table>')
			fp.write("<p>%s</p>" % data_time)
			fp.write(r'<table style="width:50%;border-collapse:collapse;"')
			title_list = ["  ",]
			title_list.extend(self.DEFAULT_SHOW_TRACK_FIELD_NAME)
			add_table_row(fp, title_list)
			for key, value in self.stock_data_dict.items():
				# value.update(record_data_dict[key])
				marker = "* " if self.__is_trailing_stop_triggered(value['啟動停利']) else "  "
				data_list = [marker, key,]
				# import pdb; pdb.set_trace()
				for field_name in self.DEFAULT_SHOW_TRACK_FIELD_NAME[1:9]:
					data_list.append(value[field_name])
				# import pdb; pdb.set_trace()
				diff_value = 0.00
				diff_value_percentage = 0.00
				if value["成交"] - value["平圴成本"] > 0:
					diff_value = self.__float(value['成交'] - value['停利價格'])
					diff_value_percentage = self.__float(diff_value / value['停利價格'] * 100.0)
				data_list.extend([diff_value, diff_value_percentage,])
				data_str_list = list(map(lambda x : str(x), data_list))
				add_table_row(fp, data_str_list)
			fp.write('</table>')
			fp.write("<hr>")

			if len(take_profit_list) != 0 or len(loss_list) != 0:
				fp.write("<div>")
				fp.write("<p>************************************************</p>")
				if len(take_profit_list) != 0:
					take_profit_string = "停利: %s" % " ".join(take_profit_list)
					fp.write("<p>%s</p>" % take_profit_string)
				if len(loss_list) != 0:
					loss_string = "虧損: %s" % " ".join(loss_list)
					fp.write("<p>%s</p>" % loss_string)
				fp.write("<p>************************************************</p>")
				fp.write("</div>")


	def print_filepath(self):
		print("************** File Path **************")
		for key, value in self.filepath_dict.items():
			if not self.__check_file_exist(value):
				print("The file[%s] does NOT exist !!!" % value)
			else:
				print("%s: %s   %s" % (key, value, self.__get_file_modification_date_str(value)))


	def output_record_file_template(self):
# 商品,平圴成本,股數,最大獲利,停利價格,啟動停利
		TEMPLATE_DATA = {
			"商品": ["2330", "2317",],
			"平圴成本": [1000.00, 150.00,],
			"股數": [100, 1000,],
		}
		line_list = []
		name_list = self.DEFAULT_RECORD_FIELD_NAME
		line_list.append(",".join(name_list))
		for index in range(2):
			data_list = [TEMPLATE_DATA[name_list[0]][index], TEMPLATE_DATA[name_list[1]][index], TEMPLATE_DATA[name_list[2]][index]]
			line_list.append(",".join(map(str, data_list)))
		template_record_filepath = self.xcfg['record_filepath'] + ".tmpl"
		with open(template_record_filepath, 'w') as fp:
			for line in line_list:
				fp.write("%s\n" % line)


	def output_customized_config_file_template(self):
		TEMPLATE_DATA = {
			"2330":{
				"trailing_stop_ratio": 0.35,
				"trigger_trailing_stop_profit_ratio": 0.2
			},
			"2317":{
				"trigger_trailing_stop_profit_ratio": 0.1
			},
			"2454":{
				"trailing_stop_ratio": 0.25,
			}
		}
		template_customized_config_filepath = self.xcfg['customized_config_filepath'] + ".tmpl"
		with open(template_customized_config_filepath, 'w') as f:
			json.dump(TEMPLATE_DATA, f, indent=4)


	@property
	def ReadFromScrapy(self):
		return self.xcfg["read_from_scrapy"]


	@ReadFromScrapy.setter
	def ReadFromScrapy(self, read_from_scrapy):
		self.xcfg["read_from_scrapy"] = read_from_scrapy


	@property
	def ForceUpdateRecord(self):
		return self.xcfg["force_update_record"]


	@ForceUpdateRecord.setter
	def ForceUpdateRecord(self, force_update_record):
		self.xcfg["force_update_record"] = force_update_record


	@property
	def MonitorMode(self):
		return self.xcfg["monitor_mode"]


	@MonitorMode.setter
	def MonitorMode(self, monitor_mode):
		self.xcfg["monitor_mode"] = monitor_mode


	@property
	def MonitorTimeInterval(self):
		return self.xcfg["monitor_time_interval"]


	@MonitorTimeInterval.setter
	def MonitorTimeInterval(self, monitor_time_interval):
		self.xcfg["monitor_time_interval"] = monitor_time_interval


	@property
	def CurTimeString(self):
		return self.__get_cur_timestr()


	@property
	def ShowResult(self):
		return self.xcfg["show_result"]


	@ShowResult.setter
	def ShowResult(self, show_result):
		self.xcfg["show_result"] = show_result


	@property
	def OutputHtmlResult(self):
		return self.xcfg["output_html_result"]


	@OutputHtmlResult.setter
	def OutputHtmlResult(self, output_html_result):
		self.xcfg["output_html_result"] = output_html_result


	@property
	def ShowScrapyProgress(self):
		return self.xcfg["show_scrapy_progress"]


	@ShowScrapyProgress.setter
	def ShowScrapyProgress(self, show_scrapy_progress):
		self.xcfg["show_scrapy_progress"] = show_scrapy_progress


	@property
	def Worksheet(self):
		return self.__get_worksheet()


	@property
	def CanTrack(self):
		# time_check = datetime.time(13, 36, 1)
		can_track = True
		if self.first_track:
			self.first_track = False
		else:
			time_check = self.__get_cur_time(True)
			can_track = self.__check_time_in_range(self.DEFAULT_CAN_SCRAPE_TIME_RANGE_START, self.DEFAULT_CAN_SCRAPE_TIME_RANGE_END, time_check)
		return can_track


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Print help')

	parser.add_argument('-t', '--track', required=False, action='store_true', help='Track specific targets.')
	parser.add_argument('--read_from_scrapy', required=False, action='store_true', help='Read stock data from scrapy. Caution: Only take effect for the "track" argument')
	parser.add_argument('--force_update_record', required=False, action='store_true', help='Update the record file forcibly. Caution: Only take effect for the "track" argument')
	parser.add_argument('-s', '--show_result', required=False, action='store_true', help='Show the tracking result of specific targets.')
	parser.add_argument('--output_html_result', required=False, action='store_true', help='Output the result in a html file')
	parser.add_argument('-m', '--monitor_mode', required=False, action='store_true', help='Monitor mode. Execute periodically')
	parser.add_argument('--monitor_time_interval', required=False, help='Time interval of monitor mode')
	parser.add_argument('--print_filepath', required=False, action='store_true', help='Print the filepaths used in the process and exit.')
	parser.add_argument('--output_record_file_template', required=False, action='store_true', help='Output a record file as a template and exit.')
	parser.add_argument('--output_customized_config_file_template', required=False, action='store_true', help='Output a customized config file as a template and exit.')
	parser.add_argument('--show_scrapy_progress', required=False, action='store_true', help='Show Scrapy progress')
	parser.add_argument('--default', required=False, action='store_true', help='Exploit the default settings: -ts --read_from_scrapy --output_html_result --show_scrapy_progress')
	args = parser.parse_args()

	cfg = {}
	# import pdb; pdb.set_trace()
	with TakeProfitTracker(cfg) as obj:
		# print("Check Scrapy: %s" % ("True" if obj.CanTrack else "False"))
		if args.print_filepath:
			obj.print_filepath()
			sys.exit(0)
		if args.output_record_file_template:
			obj.output_record_file_template()
			sys.exit(0)
		if args.output_customized_config_file_template:
			obj.output_customized_config_file_template()
			sys.exit(0)
		if args.read_from_scrapy:
			obj.ReadFromScrapy = True
		if args.force_update_record:
			obj.ForceUpdateRecord = True
		if args.show_result:
			obj.ShowResult = True
		if args.output_html_result:
			obj.OutputHtmlResult = True
		if args.monitor_mode:
			obj.MonitorMode = True
		if args.monitor_time_interval:
			obj.MonitorTimeInterval = int(args.monitor_time_interval)
		if args.show_scrapy_progress:
			obj.ShowScrapyProgress = True
		track = args.track
		if args.default:
			track = True
			obj.ReadFromScrapy = True
			obj.ShowResult = True
			obj.OutputHtmlResult = True
			obj.ShowScrapyProgress = True
		if track:
			while True:
				if obj.CanTrack:
					try:
						obj.track()
					except ScrapyError:
						pass
					if not obj.MonitorMode:
						break
					obj.refresh_data()
				time.sleep(obj.MonitorTimeInterval)
