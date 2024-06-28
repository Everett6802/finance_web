#! /usr/bin/python
# -*- coding: utf8 -*-

import os

import xlrd
# import xlsxwriter
import argparse


class TakeProfitTracker(object):

	DEFAULT_DATA_FOLDERPATH =  "C:\\停利追蹤"
	DEFAULT_SOURCE_FILENAME = "take_profile_tracker"
	DEFAULT_SOURCE_FULL_FILENAME = "%s.xlsx" % DEFAULT_SOURCE_FILENAME
	DEFAULT_RECORD_FILENAME = "take_profile_tracker_record"
	DEFAULT_RECORD_FULL_FILENAME = "%s.txt" % DEFAULT_RECORD_FILENAME
	DEFAULT_TRAILING_STOP_RATIO = 0.7

	def __init__(self, cfg):
		self.xcfg = {
			"data_folderpath": None,
			"source_filename": self.DEFAULT_SOURCE_FULL_FILENAME,
			"record_filename": self.DEFAULT_RECORD_FULL_FILENAME,
			"trailing_stop_ratio": self.DEFAULT_TRAILING_STOP_RATIO,
		}
		self.xcfg.update(cfg)
		self.xcfg["data_folderpath"] = self.DEFAULT_DATA_FOLDERPATH if self.xcfg["data_folderpath"] is None else self.xcfg["data_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FULL_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["source_filename"])
		self.xcfg["record_filename"] = self.DEFAULT_RECORD_FULL_FILENAME if self.xcfg["record_filename"] is None else self.xcfg["record_filename"]
		self.xcfg["record_filepath"] = os.path.join(self.xcfg["data_folderpath"], self.xcfg["record_filename"])

		self.workbook = None
		self.worksheet = None


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


	def __read_worksheet(self, worksheet, filterd_stock_id_list=None):
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
		line_list = self.__get_line_list_from_file(self.xcfg["record_filepath"])
		record_data = {}
		# import pdb; pdb.set_trace()
		title_list = line_list[0].split(",")
		title_list_len = len(title_list)
		for line in line_list[1:]:
			line_data_list = line.split(",")
			line_data_list_len = len(line_data_list)
			if line_data_list_len < title_list_len:
				len_diff = title_list_len - line_data_list_len
				line_data_list.extend([None,] * len_diff)
			record_data[line_data_list[0]] = dict(zip(title_list[1:], line_data_list[1:])) 
		return record_data


	def track(self):
# ['商品', '成交', '漲幅%', '漲跌']
		# import pdb; pdb.set_trace()
		record_data = self.__read_record()
		stock_data_dict = self.__read_worksheet(self.worksheet, filterd_stock_id_list=record_data.keys())
# update() doesn't return any value (returns None).
		# stock_data_dict = [(key, value, record_data[key], value.update(record_data[key])) for key, value in stock_data_dict.items()]
		# stock_data_dict.update(record_data)
		for key, value in stock_data_dict.items():
			value.update(record_data[key])
			if value["成交"] - value["平圴成本"] > 0:
				profile = (value["成交"] - value["平圴成本"]) * value["股數"]
				if value["最大獲利"] is None or profile > data["最大獲利"]:
					value["最大獲利"] = profile
	 				value["停利價格"] = profile * self.xcfg["trailing_stop_ratio"] / value["股數"] + value["平圴成本"]
				else:
					if value["成交"] < value["停利價格"]:
						print("停利: %s" % key)
		# print(stock_data_dict)


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Print help')
	
	cfg = {}
	
	with TakeProfitTracker(cfg) as obj:
		obj.track()
