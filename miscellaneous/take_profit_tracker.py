#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import re
import xlrd
# import xlsxwriter
import argparse


class TakeProfitTracker(object):

	DEFAULT_DATA_FOLDERPATH =  "C:\\停利追蹤"
	DEFAULT_SOURCE_FILENAME = "take_profile_tracker"
	DEFAULT_SOURCE_FULL_FILENAME = "%s.xlsx" % DEFAULT_SOURCE_FILENAME
	DEFAULT_RECORD_FILENAME = "take_profile_tracker_record"
	DEFAULT_RECORD_FULL_FILENAME = "%s.txt" % DEFAULT_RECORD_FILENAME
	DEFAULT_STOCK_SYMBOL_LOOKUP_FILENAME = "股號查詢"
	DEFAULT_STOCK_SYMBOL_LOOKUP_FULL_FILENAME = "%s.xlsx" % DEFAULT_STOCK_SYMBOL_LOOKUP_FILENAME
	DEFAULT_TRAILING_STOP_RATIO = 0.7
# 代碼,平圴成本,股數,最大獲利,停利價格
	DEFAULT_RECORD_FIELD_NAME = ["代碼", "平圴成本", "股數", "最大獲利", "停利價格"]
	DEFAULT_RECORD_FIELD_TYPE = [str, float, int, int, float]
	YAHOO_STOCK_URL_FORMAT = "https://tw.stock.yahoo.com/quote/%s.TW"

	def __init__(self, cfg):
		self.xcfg = {
			"data_folderpath": None,
			"source_filename": self.DEFAULT_SOURCE_FULL_FILENAME,
			"record_filename": self.DEFAULT_RECORD_FULL_FILENAME,
			"stock_symbol_lookup_filename": self.DEFAULT_STOCK_SYMBOL_LOOKUP_FULL_FILENAME,
			"trailing_stop_ratio": self.DEFAULT_TRAILING_STOP_RATIO,
		}
		self.xcfg.update(cfg)
		self.xcfg["data_folderpath"] = self.DEFAULT_DATA_FOLDERPATH if self.xcfg["data_folderpath"] is None else self.xcfg["data_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FULL_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["source_filename"])
		self.xcfg["record_filename"] = self.DEFAULT_RECORD_FULL_FILENAME if self.xcfg["record_filename"] is None else self.xcfg["record_filename"]
		self.xcfg["record_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["record_filename"])
		self.xcfg["stock_symbol_lookup_filename"] = self.DEFAULT_STOCK_SYMBOL_LOOKUP_FULL_FILENAME if self.xcfg["stock_symbol_lookup_filename"] is None else self.xcfg["stock_symbol_lookup_filename"]
		self.xcfg["stock_symbol_lookup_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["stock_symbol_lookup_filename"])

		self.workbook = None
		self.worksheet = None
		self.can_lookup_stock_symbol = False
		self.stock_symbol_lookup_dict = None  # 股名 -> 股號
		self.stock_symbol_reverse_lookup_dict = None  # 股號 -> 股名
		self.can_scrape = self.__can_scrape()
		self.requests_module = None
		self.beautifulsoup_class = None


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
	def __get_line_list_from_file(self, filepath):
		# import pdb; pdb.set_trace()
		if not self.__check_file_exist(filepath):
			raise RuntimeError("The file[%s] does NOT exist" % filepath)
		line_list = []
		with open(filepath, 'r') as fp:
			for line in fp:
				if line.startswith("#"): continue
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


	def __enter__(self):
# Open the workbook
		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		self.worksheet = self.workbook.sheet_by_index(0)
		return self


	def __exit__(self, type, msg, traceback):
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
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


	def __read_worksheet(self, worksheet, filterd_stock_id_list=None):
# Check if it's required to transform from stock name to stock symbol
		worksheet_data = {}
		# import pdb; pdb.set_trace()
		title_list = []
# title
		for column_index in range(1, worksheet.ncols):
			title_value = worksheet.cell_value(0, column_index)
			title_list.append(title_value)
		# print(title_list)
		# import pdb; pdb.set_trace()
		need_lookup_stock_symbol = False
		if re.match("[\d]{4,}", worksheet.cell_value(1, 0)) is None:
			if not self.can_lookup_stock_symbol:
				raise RuntimeError("No stock symbol lookup table !!!")
			else:
				need_lookup_stock_symbol = True
# data
		for row_index in range(1, worksheet.nrows):
			data_key = worksheet.cell_value(row_index, 0)
			if need_lookup_stock_symbol:
				data_key = self.stock_symbol_lookup_dict[data_key]
			if (filterd_stock_id_list is not None) and (data_key not in filterd_stock_id_list):
				continue
			data_list = []
			for column_index in range(1, worksheet.ncols):
				data_value = worksheet.cell_value(row_index, column_index)
				data_list.append(data_value)
			# print("%s: %s" % (data_key, data_list))
			data_dict = dict(zip(title_list, data_list))
			worksheet_data[data_key] = data_dict
		return worksheet_data


	def __read_record(self):
# 代碼,平圴成本,股數,最大獲利,停利價格
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
			record_data_dict[line_data_list[0]] = dict(zip(title_list[1:], line_data_list[1:])) 
		return record_data_dict


	def __write_record(self, record_data_dict):
# 代碼,平圴成本,股數,最大獲利,停利價格
		# import pdb; pdb.set_trace()
		with open(self.xcfg['record_filepath'], 'w') as fp:
			line = ",".join(self.DEFAULT_RECORD_FIELD_NAME)
			fp.write("%s\n" % line)
			for stock_symbol in record_data_dict.keys():
				line_data_list = [stock_symbol,]
				line_data_list.append(record_data_dict[stock_symbol]["平圴成本"])
				line_data_list.append(record_data_dict[stock_symbol]["股數"])
				line_data_list.append(record_data_dict[stock_symbol]["最大獲利"])
				line_data_list.append(record_data_dict[stock_symbol]["停利價格"])
				line_data_list = map(str, line_data_list)
				line = ",".join(line_data_list)
				fp.write("%s\n" % line)


	def __read_stock_symbol_mapping_table(self):
# 代碼,商品,股本
		if self.can_lookup_stock_symbol: return
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
		import pdb; pdb.set_trace()
		url = self.YAHOO_STOCK_URL_FORMAT % stock_symbol
		resp = self.__get_requests_module().get(url)
		if re.search(stock_symbol, resp.text) is None:
			raise ValueError("The stock[%s] does NOT exist" % stock_symbol)
		# print(resp.text)
		beautifulsoup_class = self.__get_beautifulsoup_class()
		soup = beautifulsoup_class(resp.text, "html.parser")
		div = soup.find("div", {"id": "main-2-QuoteOverview-Proxy"})
		div1 = div.find_all("div", {"class": "D(f)"})
		div2 = div1.find_all("div", {"class": "Pos(r)"})
		lis = div2.find_all("li")
		for li in lis:
			print(li.text)
# 		try:
# 			table_trs = table[0].find_all("tr")
# 		except Exception as e:
# # Too many query requests from your ip, please wait and try again later!!
# 			print(e)
# 			# raise RetryException("Too many query requests from your ip, please wait and try again later")


	def scrape(self):
		self.__scrape_stock_price("2317")


	def track(self):
# ['商品', '成交', '漲跌', '漲幅%']
		record_data_dict = self.__read_record()
		stock_data_dict = self.__read_worksheet(self.worksheet, filterd_stock_id_list=record_data_dict.keys())			
# update() doesn't return any value (returns None).
		# stock_data_dict = [(key, value, record_data_dict[key], value.update(record_data_dict[key])) for key, value in stock_data_dict.items()]
		# stock_data_dict.update(record_data_dict)
		need_update_record = False
		# import pdb; pdb.set_trace()
		for key, value in stock_data_dict.items():
			value.update(record_data_dict[key])
			if value["成交"] - value["平圴成本"] > 0:
				# import pdb; pdb.set_trace()
				profile = int((value["成交"] - value["平圴成本"]) * value["股數"])
				if value["最大獲利"] is None or profile > value["最大獲利"]:
					need_update_record = True
					value["最大獲利"] = profile
					tmp = profile * self.xcfg["trailing_stop_ratio"] / value["股數"] + value["平圴成本"]
					value["停利價格"] = float("%.2f" % tmp)
				else:
					if value["成交"] < value["停利價格"]:
						print("停利: %s" % key)
			else:
				if value["最大獲利"] is None:
					value["最大獲利"] = 0
					value["停利價格"] = 0.00
					need_update_record = True
		if need_update_record:
			self.__write_record(stock_data_dict)
		# print(stock_data_dict)


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Print help')
	
	cfg = {}
	
	with TakeProfitTracker(cfg) as obj:
		obj.scrape()
