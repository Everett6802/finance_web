#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import xlrd
import argparse
import errno
import math
import locale
import datetime
import statistics
from collections import OrderedDict


def write_to_file(func):
	def wrapper(obj, *args, **kwargs):
		obj.redirect_stdout2file()
		func(obj, *args, **kwargs)
		obj.redirect_file2stdout()
	return wrapper
		

class StockFluctuationStatistics(object):

	@staticmethod
	def __os_is_chinese():
		locale_module_member_list = dir(locale)
		encoding = None
		if 'getdefaultlocale' in locale_module_member_list:
			_, encoding = locale.getdefaultlocale()
		elif 'getencoding' in locale_module_member_list:
			encoding = locale.getencoding()
		if encoding is None:
			raise RuntimeError("Fails to find the OS encoding")
		return True if (encoding == "cp950") else False


	@staticmethod
	def get_correlation(list_x, list_y):
		list_x_len = len(list_x)
		list_y_len = len(list_y)
		if list_x_len != list_y_len: raise ValueError("The lengthes of the 2 lists are NOT identical: %d, %d" % (list_x_len, list_y_len))
		list_x_mean = statistics.mean(list_x)
		list_y_mean = statistics.mean(list_y)
		# list_x_std = statistics.stdev(list_x)
		# list_y_std = statistics.stdev(list_y)
		sum_xy = 0
		sum_xx = 0
		sum_yy = 0
		# import pdb; pdb.set_trace()
		for index in range(list_x_len):
			sum_xy += (list_x[index] - list_x_mean) * (list_y[index] - list_y_mean)
			sum_xx += pow((list_x[index] - list_x_mean), 2)
			sum_yy += pow((list_y[index] - list_y_mean), 2)
		# import pdb; pdb.set_trace()
		return sum_xy / math.sqrt(sum_xx * sum_yy)


	@classmethod
	def check_file_exist(cls, filepath):
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
	def get_line_list_from_file(cls, filepath, startswith=None):
		# import pdb; pdb.set_trace()
		if not cls.check_file_exist(filepath):
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
	def parse_statistics_data_from_file(cls, filepath):
		line_list = cls.get_line_list_from_file(filepath)
		statistics_data_dict = OrderedDict()
		for line in line_list:
			elem_list = line.split()
			assert len(elem_list) == 2, "The length(1) of elem_list[%s] should be 2" % elem_list
			[key, value_tmp] = elem_list
			elem_list = value_tmp.split("[")
			assert len(elem_list) == 2, "The length(2) of elem_list[%s] should be 2" % elem_list
			[value, rest] = elem_list
			statistics_data_dict[key] = float(value)
		return statistics_data_dict


	DEFAULT_SOURCE_FILENAME = "加權指數歷史資料2000-2024.xlsx"
	DEFAULT_SOURCE_FILENAME2 = "期貨指數歷史資料2000-2024.xlsx"
	# DEFAULT_SOURCE_FULL_FILENAME = "%s.xlsx" % DEFAULT_SOURCE_FILENAME
	DEFAULT_TIME_FIELD_NAME = "時間"	
	DEFAULT_CLOSING_PRICE_FIELD_NAME = "收盤價"	
	DEFAULT_DATE_BASE_NUMBER = 36526
	DEFAULT_DATE_BASE = datetime.date(2000, 1, 1)
	DEFAULT_TRADE_DATE_IS_HOLIDAY_FILENAME = "trade_date_is_holiday"
	DEFAULT_RISE_PERCENTAGE_THRESHOLD = 80.0
	DEFAULT_FALL_PERCENTAGE_THRESHOLD = 20.0
	DEFAULT_OUTPUT_RESULT_FILENAME = "stock_fluctuation_statistics.txt"

	@classmethod
	def __get_google_cloud_root_foldername(cls):
		return "我的雲端硬碟" if cls.__os_is_chinese() else "My Drive"


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
	def __is_leap_year(cls, year):
		return True if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0) else False


	def __init__(self, cfg):
		self.xcfg = {
			"data_folderpath": None,
			"source_filename": None,
			"trade_date_is_holiday_filename": self.DEFAULT_TRADE_DATE_IS_HOLIDAY_FILENAME,
			"trade_date_is_holiday_folderpath": None,
			"trade_date_string": None,
			"rise_percentage_threshold": self.DEFAULT_RISE_PERCENTAGE_THRESHOLD,
			"fall_percentage_threshold": self.DEFAULT_FALL_PERCENTAGE_THRESHOLD,
			"output_result_filename": self.DEFAULT_OUTPUT_RESULT_FILENAME,
			"statistics_dependency_filename1": self.DEFAULT_SOURCE_FILENAME,
			"statistics_dependency_filename2": self.DEFAULT_SOURCE_FILENAME2,
		}
		self.xcfg.update(cfg)
		# import pdb; pdb.set_trace()
		self.DEFAULT_DATA_FOLDERPATH =  "G:\\{root_foldername}\\數據".format(root_foldername=self.__get_google_cloud_root_foldername())
		self.DEFAULT_OUTPUT_FOLDERPATH =  self.DEFAULT_DATA_FOLDERPATH  # "C:\\Users\\%s" % os.getlogin()
	
		self.xcfg["data_folderpath"] = self.DEFAULT_DATA_FOLDERPATH if self.xcfg["data_folderpath"] is None else self.xcfg["data_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		# source_full_filename = "%s.xlsx" % self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["source_filename"])  # source_full_filename)
		self.xcfg["trade_date_is_holiday_filename"] = self.DEFAULT_TRADE_DATE_IS_HOLIDAY_FILENAME if self.xcfg["trade_date_is_holiday_filename"] is None else self.xcfg["trade_date_is_holiday_filename"]
		# self.xcfg["trade_date_is_holiday_folderpath"] = os.getcwd() if self.xcfg["trade_date_is_holiday_folderpath"] is None else self.xcfg["trade_date_is_holiday_folderpath"]
		self.xcfg["trade_date_is_holiday_folderpath"] = self.DEFAULT_OUTPUT_FOLDERPATH if self.xcfg["trade_date_is_holiday_folderpath"] is None else self.xcfg["trade_date_is_holiday_folderpath"]
		self.xcfg["trade_date_is_holiday_filepath"] = os.path.join(self.xcfg["trade_date_is_holiday_folderpath"], self.xcfg["trade_date_is_holiday_filename"])
		# self.xcfg["output_result_filepath"] = os.path.join(os.getcwd(), self.xcfg["output_result_filename"])
		self.xcfg["output_result_filepath"] = os.path.join(self.DEFAULT_OUTPUT_FOLDERPATH, self.xcfg["output_result_filename"])
		self.xcfg["statistics_dependency_filename1"] = self.DEFAULT_SOURCE_FILENAME if self.xcfg["statistics_dependency_filename1"] is None else self.xcfg["statistics_dependency_filename1"]
		self.xcfg["statistics_dependency_filename2"] = self.DEFAULT_SOURCE_FILENAME2 if self.xcfg["statistics_dependency_filename2"] is None else self.xcfg["statistics_dependency_filename2"]
		self.xcfg["statistics_dependency_filepath1"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["statistics_dependency_filename1"])
		self.xcfg["statistics_dependency_filepath2"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["statistics_dependency_filename2"])

		self.filepath_dict = OrderedDict()
		self.filepath_dict["source_filepath"] = self.xcfg["source_filepath"]
		self.filepath_dict["trade_date_is_holiday_filepath"] = self.xcfg["trade_date_is_holiday_filepath"]
		self.filepath_dict["output_result_filepath"] = self.xcfg["output_result_filepath"]
		self.filepath_dict["statistics_dependency_filepath1"] = self.xcfg["statistics_dependency_filepath1"]
		self.filepath_dict["statistics_dependency_filepath2"] = self.xcfg["statistics_dependency_filepath2"]

		self.workbook = None
		self.trade_date_is_holiday_date_list = None
		self.cur_year = datetime.datetime.now().year
		self.is_leap_year = self.__is_leap_year(self.cur_year)
		self.trade_opportunity_data_list = None
		self.output_result_file = None
		self.stdout_tmp = None


	def __enter__(self):
		return self


	def __exit__(self, type, msg, traceback):
		if self.output_result_file is not None:
			self.output_result_file.close()
			self.output_result_file = None
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


	def __date_str2obj(self, date_str):
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
		date_obj = None
		try:
			date_obj = datetime.date(year, month, day)
		except Exception as e:
			print("Unsupport date format[%s] due to: %s" % (date_str, str(e)))
			# import pdb; pdb.set_trace()
			raise e
		return date_obj


	def __get_worksheet(self):
		if self.workbook is None:
			# import pdb; pdb.set_trace()
			if not self.__check_file_exist(self.xcfg["source_filepath"]):
				raise RuntimeError("The worksheet[%s] does NOT exist" % self.xcfg["source_filepath"])
			self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
			self.worksheet = self.workbook.sheet_by_index(0)
		return self.worksheet


	def __read_worksheet(self):
# Check if it's required to transform from stock name to stock symbol
		worksheet_data = {}
		# import pdb; pdb.set_trace()
		title_list = []
		data_list = []
# title
		for column_index in range(0, self.Worksheet.ncols):
			title_value = self.Worksheet.cell_value(0, column_index)
			title_list.append(title_value)
		# print(title_list)
		# import pdb; pdb.set_trace()
# data
		for row_index in range(1, self.Worksheet.nrows):
			entry_list = []
			can_add = True
			for column_index in range(0, self.Worksheet.ncols):
				entry_value = self.Worksheet.cell_value(row_index, column_index)
				if column_index == 0:
					day_diff = int(entry_value) - self.DEFAULT_DATE_BASE_NUMBER
					entry_date = self.DEFAULT_DATE_BASE + datetime.timedelta(days=day_diff)
					if entry_date < datetime.date(2010, 1, 1):
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


	def __get_trade_date_is_holiday_date_list(self):
		if self.trade_date_is_holiday_date_list is None:
			if not self.__check_file_exist(self.xcfg["trade_date_is_holiday_filepath"]):
				raise RuntimeError("The file of date list[%s] does NOT exist" % self.xcfg["trade_date_is_holiday_filepath"])
			self.trade_date_is_holiday_date_list = []
			# cur_date = datetime.datetime.now()
			# cur_year = cur_date.year
			# import pdb; pdb.set_trace()
			with open(self.xcfg["trade_date_is_holiday_filepath"], "r") as fp:
				for line in fp:
					holiday_date = datetime.datetime.strptime(line.strip("\n"), "%m/%d")
					self.trade_date_is_holiday_date_list.append(datetime.date(self.cur_year, holiday_date.month, holiday_date.day))
		return self.trade_date_is_holiday_date_list


	def __calculate_historical_fluctuation(self):
		worksheet_data = self.__read_worksheet()
		time_index = worksheet_data["title"].index(self.DEFAULT_TIME_FIELD_NAME)
		closing_price_index = worksheet_data["title"].index(self.DEFAULT_CLOSING_PRICE_FIELD_NAME)
		data_len = len(worksheet_data["data"])
		fluctuation_data = {}
		# import pdb; pdb.set_trace()
		for index in range(1, data_len):
			data_time = worksheet_data["data"][index][time_index]
			data_fluctuation = worksheet_data["data"][index][closing_price_index] - worksheet_data["data"][index - 1][closing_price_index]
			date_time_date = data_time[:data_time.rindex("/")]
			if date_time_date not in fluctuation_data:
				fluctuation_data[date_time_date] = []
			fluctuation_data[date_time_date].append(data_fluctuation)
		# import pdb; pdb.set_trace()
		fluctuation_data = OrderedDict(list(sorted(fluctuation_data.items(), key=lambda x: x[0])))
		return fluctuation_data


	def __check_date_can_trade(self, check_date):
		if check_date.weekday() not in [5, 6,]: # Saturday and Sunday
			if check_date not in self.__get_trade_date_is_holiday_date_list():
				return True
		return False


	def __get_trade_date(self, check_start_date):
		count = 1
		while True:
			trade_date = check_start_date - datetime.timedelta(days=count)
			if self.__check_date_can_trade(trade_date):
				return trade_date
			if count >= 365:
				break
			count += 1
		raise RuntimeError("Fail to find the trade date from the date: %s" % check_start_date)	


	def __analyze_historical_fluctuation(self, silent=True):
		fluctuation_data = self.__calculate_historical_fluctuation()
		trade_opportunity_data_list = []
		for key, value in fluctuation_data.items():
			data_len = len(value)
			value_rise = list(filter(lambda x: x > 0, value))
			data_rise_len = len(value_rise)
			value_fall = list(filter(lambda x: x < 0, value))
			# data_fall_len = len(value_fall)
			data_rise_percentage = round(float(data_rise_len) * 100.0 / data_len, 1)
			if self.xcfg["rise_percentage_threshold"] is not None and data_rise_percentage >= self.xcfg["rise_percentage_threshold"]:
				if not silent: print("%s  %.1f[%d/%d]  %.1f  %.2f  %.1f  %.2f" % (key, data_rise_percentage, data_rise_len, data_len, statistics.mean(value), statistics.stdev(value), statistics.mean(value_rise), statistics.stdev(value_rise)))
				if key == "02/29" and not self.is_leap_year: continue
				trade_opportunity_data = {
					"trade_opportunity_date": self.__date_str2obj(key),
					"rise_or_fall": "rise",
					"historic_value": value,
					"historic_value_rise": value_rise
				}
				trade_opportunity_data_list.append(trade_opportunity_data)
			elif self.xcfg["fall_percentage_threshold"] is not None and data_rise_percentage <= self.xcfg["fall_percentage_threshold"]:
				if not silent: print("%s  %.1f[%d/%d]  %.1f  %.2f  %.1f  %.2f" % (key, data_rise_percentage, data_rise_len, data_len, statistics.mean(value), statistics.stdev(value), statistics.mean(value_fall), statistics.stdev(value_fall)))
				if key == "02/29" and not self.is_leap_year: continue
				trade_opportunity_data = {
					"trade_opportunity_date": self.__date_str2obj(key),
					"rise_or_fall": "fall",
					"historic_value": value,
					"historic_value_fall": value_fall
				}
				trade_opportunity_data_list.append(trade_opportunity_data)
		return trade_opportunity_data_list


	def __get_trade_opportunity_data_list(self, silent=True):
		if self.trade_opportunity_data_list is None:
			self.trade_opportunity_data_list = self.__analyze_historical_fluctuation(silent)
			for trade_opportunity_data in self.trade_opportunity_data_list:
				trade_opportunity_data["trade_date"] = None
				check_date = trade_opportunity_data["trade_opportunity_date"]
				if not self.__check_date_can_trade(check_date):
					# if not silent: print("The date[%s] is NOT a trade date" % check_date)
					continue
				# import pdb; pdb.set_trace()
				trade_opportunity_data["trade_date"] = self.__get_trade_date(check_date)
				if not silent: 
					print("%s -> %s:%s" % (check_date, ("X" if trade_opportunity_data["trade_date"] is None else trade_opportunity_data["trade_date"]), ("Bull" if trade_opportunity_data["rise_or_fall"] == "rise" else "Bear")))
		return self.trade_opportunity_data_list


	def redirect_stdout2file(self):
		# import pdb; pdb.set_trace()
# The output is now directed to the file
		if self.output_result_file is None:
			self.output_result_file = open(self.xcfg["output_result_filepath"], 'w')
# Store the current STDOUT object for later use
		self.stdout_tmp = sys.stdout
# Redirect STDOUT to the file
		sys.stdout = self.output_result_file


	def redirect_file2stdout(self):
# Restore the original STDOUT
		sys.stdout = self.stdout_tmp
# Close the file handle
		if self.output_result_file is not None:
			self.output_result_file.close()
			self.output_result_file = None


	def get_check_date_trade_info(self, check_date=None):
		if check_date is None: check_date = datetime.datetime.now().date()
		# trade_opportunity_data_list = self.find_trade_date()
		# import pdb; pdb.set_trace()
		trade_date_list = [trade_opportunity_data["trade_date"] for trade_opportunity_data in self.__get_trade_opportunity_data_list()]
		trade_date_list_len = len(trade_date_list)
		if trade_date_list_len == 0:
			raise RuntimeError("No trade date !!!")
		try:
			index = trade_date_list.index(check_date)
			return self.__get_trade_opportunity_data_list()[index]
		except:
			pass
		return None


	def get_latest_date_trade_info(self, check_date=None):
		if check_date is None: check_date = datetime.datetime.now().date()
		# import pdb; pdb.set_trace()
		trade_date_list = [trade_opportunity_data["trade_date"] for trade_opportunity_data in self.__get_trade_opportunity_data_list()]
		trade_date_list_len = len(trade_date_list)
		if trade_date_list_len == 0:
			raise RuntimeError("No trade date !!!")
		# import pdb; pdb.set_trace()
		for index, trade_date in enumerate(trade_date_list):
			if trade_date is None: continue
			if trade_date >= check_date:
				return self.__get_trade_opportunity_data_list()[index] 
		# filtered_date_list = list(filter(lambda x : x >= check_date, trade_date_list))
		# filtered_date_list_len = len(filtered_date_list)
		# if filtered_date_list_len != 0:
		# 	index = trade_date_list_len - filtered_date_list_len
		# 	return self.__get_trade_opportunity_data_list()[index] 
		return None


	def parse_check_date_trade_info(self, check_date=None):
		if check_date is None: check_date = datetime.datetime.now().date()
		# trade_opportunity_data = self.get_check_date_trade_info(check_date)
		trade_opportunity_data = self.get_latest_date_trade_info(check_date)
		# import pdb; pdb.set_trace()
		if trade_opportunity_data is None:
			print("No trade opportunity in %d" % self.cur_year)
		else:
			if trade_opportunity_data["trade_date"] != check_date:
				print("%s -> Not a trade day" % check_date)
				print("================================\nThe next trade date: ")
			rise_or_fall = trade_opportunity_data["rise_or_fall"]
			print("%s -> Go %s" % (trade_opportunity_data["trade_date"], "Long" if rise_or_fall == "rise" else "Short"))
			print("Trade opportunity date: %s" % trade_opportunity_data["trade_opportunity_date"])
			print("Total history:")
			# import pdb; pdb.set_trace()
			historic_value = trade_opportunity_data["historic_value"]
			print("%s" % ", ".join(map(lambda x : "%.2f" % x, historic_value)))
			print("mean: %.1f  STD: %.1f" % (statistics.mean(historic_value), statistics.stdev(historic_value)))
			print("Total %s history:" % ("rise" if rise_or_fall == "rise" else "fall"))
			historic_rise_or_fall_value = trade_opportunity_data["historic_value_rise"] if rise_or_fall == "rise" else trade_opportunity_data["historic_value_fall"]
			print("%s" % ", ".join(map(lambda x : "%.2f" % x, historic_rise_or_fall_value)))
			print("mean: %.1f  STD: %.1f" % (statistics.mean(historic_rise_or_fall_value), statistics.stdev(historic_rise_or_fall_value)))
			print("probability: %.1f (%d/%d)" % (len(historic_rise_or_fall_value) * 100.0 / len(historic_value), len(historic_rise_or_fall_value), len(historic_value)))


	def check_trade_date(self, check_date=None):
		trade_opportunity_data = self.get_check_date_trade_info(check_date)
		return True if trade_opportunity_data is not None else False


	def check_trade(self):
		check_date = None
		# import pdb; pdb.set_trace()
		if self.xcfg["trade_date_string"] is not None:
			check_date = self.__date_str2obj(self.xcfg["trade_date_string"])
		self.parse_check_date_trade_info(check_date)


	def list_trade_opportunity(self):
		self.__get_trade_opportunity_data_list(False)


	@write_to_file
	def list_trade_opportunity_to_file(self):
		self.list_trade_opportunity()


	def show_statistics(self):
		fluctuation_data = self.__calculate_historical_fluctuation()
		for key, value in fluctuation_data.items():
			data_len = len(value)
			value_rise = list(filter(lambda x: x > 0, value))
			data_rise_len = len(value_rise)
			data_rise_percentage = round(float(data_rise_len) * 100.0 / data_len, 1)
			print("%s  %.1f[%d/%d]" % (key, data_rise_percentage, data_rise_len, data_len))


	@write_to_file
	def show_statistics_to_file(self):
		self.show_statistics()


	@property
	def Worksheet(self):
		return self.__get_worksheet()


	@property
	def OutputResultFilepath(self):
		return self.xcfg["output_result_filepath"]


	def print_filepath(self):
		print("************** File Path **************")
		for key, value in self.filepath_dict.items():
			print("%s: %s" % (key, value))


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Print help', formatter_class=argparse.RawTextHelpFormatter)
	parser.add_argument('-c', '--check_trade', required=False, action='store_true', help='Check the trade opportunity on a specific date and exit.')
	parser.add_argument('--trade_date', required=False,
		 help='''The trade date for checking the trade opportunity.
  Format: YYYY-mm-dd   Ex: 2025-03-11
  Format: m(m)/d(d)/YYYY   Ex: 3/11/2025
  Format: mm-dd   Ex: 03-11   Note: use current year if the year is NOT set
  Format: m(m)/d(d)   Ex: 3/11   Note: use current year if the year is NOT set''')
	# parser.add_argument('--tracked_stock_list', required=False, help='The list of specific stock targets to be trackeded.')
	parser.add_argument('-l', '--list_trade_opportunity', required=False, action='store_true', help='List trade opportunities and exit.')
	parser.add_argument('-s', '--show_statistics', required=False, action='store_true', help='Show all the statistics data and exit.')
	parser.add_argument('--source_filename', required=False, help='Set the source filename')
	parser.add_argument('-o', '--output_result', required=False, action='store_true', help='Output the result to the file')
	parser.add_argument('--output_result_filename', required=False, help='The filename of outputting the result')
	parser.add_argument('-f', '--find_statistics_dependency', required=False, action='store_true', help='Find the 2 data dependency')
	parser.add_argument('--statistics_dependency_filename_list', required=False, help='The 2 filenames of finding the data dependency. The filenames are split by comma')
	parser.add_argument('--print_filepath', required=False, action='store_true', help='Print the filepaths used in the process and exit.')
	args = parser.parse_args()

	cfg = {}
	if args.trade_date:
		cfg['trade_date_string'] = args.trade_date
	if args.source_filename:
		cfg['source_filename'] = args.source_filename
	if args.output_result_filename:
		cfg['output_result_filename'] = args.output_result_filename

	if args.find_statistics_dependency:
		statistics_data_dict1 = None
		statistics_data_dict2 = None
		statistics_dependency_filename1 = None
		statistics_dependency_filename2 = None
		if args.statistics_dependency_filename_list is not None:
			elem_list = args.statistics_dependency_filename_list.split(",")
			if len(elem_list) != 2:
				raise ValueError("The length of statistics_dependency_filename_list should be 2")
			[statistics_dependency_filename1, statistics_dependency_filename2] = elem_list
# statistics_dependency_filename1
		cfg1 = {
			"source_filename": statistics_dependency_filename1 if (statistics_dependency_filename1 is not None) else None,
			'output_result_filename': "1.tmp",
		}
		with StockFluctuationStatistics(cfg1) as obj1:
			obj1.show_statistics_to_file()
			statistics_data_dict1 = obj1.parse_statistics_data_from_file(obj1.OutputResultFilepath)
			os.remove(obj1.OutputResultFilepath)
# statistics_dependency_filename2
		cfg2 = {
			"source_filename": statistics_dependency_filename2 if (statistics_dependency_filename2 is not None) else None,
			'output_result_filename': "2.tmp",
		}
		with StockFluctuationStatistics(cfg2) as obj2:
			obj2.show_statistics_to_file()
			statistics_data_dict2 = obj2.parse_statistics_data_from_file(obj2.OutputResultFilepath)
			os.remove(obj2.OutputResultFilepath)
		assert statistics_data_dict1 is not None, "statistics_data_dict1 should not be None"
		assert statistics_data_dict2 is not None, "statistics_data_dict2 should not be None"
		if statistics_data_dict1.keys() != statistics_data_dict2.keys():
			raise ValueError("The dates of the 2 data are NOT identical")
		'''
In Python 3, dict.values() returns a dict_values object, which is a view object providing a dynamic representation 
of the dictionary's values. This differs from Python 2, where it returned a list. The dict_values object offers several 
advantages and characteristics such as the followings:
 1. Dynamic View: It reflects real-time changes to the dictionary without creating a separate copy of the values.
 2. Memory Efficiency: As a view, it doesn't create a new list of values, saving memory.
 3. No Indexing or Slicing: Unlike lists, dict_values objects don't support indexing or slicing operations.

 If you need list-like functionality, you can convert a dict_values object to a list
		'''
		correlation_value = StockFluctuationStatistics.get_correlation(list(statistics_data_dict1.values()), list(statistics_data_dict2.values()))
		print("Correlation: %.2f" % correlation_value)
	else:
		with StockFluctuationStatistics(cfg) as obj:
			if args.check_trade:
				obj.check_trade()
				sys.exit(0)
			if args.list_trade_opportunity:
				if args.output_result:
					obj.list_trade_opportunity_to_file()
				else:
					obj.list_trade_opportunity()
				sys.exit(0)
			if args.show_statistics:
				if args.output_result:
					obj.show_statistics_to_file()
				else:
					obj.show_statistics()
				sys.exit(0)
			if args.print_filepath:
				obj.print_filepath()
				sys.exit(0)
		# obj.analyze_historical_fluctuation()
		# obj.find_trade_date(silent=False)
		# print(obj.check_trade_date(datetime.date(2025,3,17)))
		# obj.parse_check_date_trade_info(datetime.date(2025,3,17))
		# test_datetime = datetime.datetime.now() - datetime.timedelta(days = 3)
		# print(test_datetime.date(), obj.check_date_can_trade(test_datetime.date()))
	# print(os.getcwd())
	# entry_date = "3/3"
	# entry_value = datetime.datetime.strptime(entry_date, "%m/%d")
	# my_date = datetime.date(2025, entry_value.month, entry_value.day)
	# print(my_date)
	# print(my_date.weekday())
	# print(my_date.date() == datetime.datetime.now().date())
