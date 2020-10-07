#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import requests
import xlrd
import argparse
import time
import json
import datetime
from lib.collector_timer import CollectorTimer


DEF_START_TIME_STR = "08:30"
DEF_END_TIME_STR = "13:46"
DEF_TIME_INTERVAL_IN_SEC = 300

class DayTradeMonitorCollector(object):

	DEFAULT_SOURCE_FILENAME = "xq_day_trade_monitor.xlsx"
	DEFAULT_SHEET_INDEX = 0

	DATA_CELL_COLUMN_TITLE_NAME_LIST = ["商品", "成交比重%", "漲幅%", "昨量", "估計量", "總量", "漲跌", "成交", "累買成筆", "累賣成筆",]
	DATA_CELL_COLUMN_TYPE_LIST = ["string", "float", "float", "int", "int", "int", "float", "float", "int", "int",]
	DATA_CELL_ROW_TITLE_NAME_LIST = ["商品", "台股指數近月(一般)", "加權指數", "台積電", "鴻海", "聯發科", "大立光", "SGX摩台近月", "富邦VIX", "恐慌指數",]

	DATA_CELL_COLUMN_DEF_START_INDEX = 1
	DATA_CELL_COLUMN_DEF_END_INDEX = 8
	DATA_CELL_ROW_COLUMN_INDEX_DICT = {
		1: [2, 3, 5, 6, 7, 8, 9,], # 台股指數近月(一般) => index list
		2: {"from": 2, "to": 8}, # 加權指數 => index range
		7: [2, 3, 5, 6, 7,], # SGX摩台近月 => index list
		# 8: {"from": 2, "to": 10}, # 加權指數 => index range
		9: [2, 6, 7,], # 恐慌指數 => index list
	}
	DATA_CELL_ROW_START_INDEX = 1
	DATA_CELL_ROW_END_INDEX = len(DATA_CELL_ROW_TITLE_NAME_LIST)
	HTTP_SUCCESS_CODE_LIST = [200, 201, 204,]

	DATETIME_FORMAT_STR = "%Y-%m-%d %H:%M:%S"
	DATE_FORMAT_STR = "%Y-%m-%d"

	@classmethod
	def __is_string(cls, value):
		is_string = False
		try:
			int(value)
		except ValueError:
			is_string = True
		return is_string


	class DateEncoder(json.JSONEncoder):  
		def default(self, obj):  
			if isinstance(obj, datetime.datetime):  
				return obj.strftime(DayTradeMonitorCollector.DATETIME_FORMAT_STR)  
			elif isinstance(obj, date):  
				return obj.strftime(DayTradeMonitorCollector.DATE_FORMAT_STR)  
			else:  
				return json.JSONEncoder.default(self, obj) 


	def __init__(self, cfg):
		self.xcfg = {
			"source_filepath": os.path.join(os.getcwd(), self.DEFAULT_SOURCE_FILENAME),
			"sheet_index": self.DEFAULT_SHEET_INDEX,
			"target_server_address": "10.206.24.219",
			"target_server_port": 5998,
			"start_time": "08:30",
			"end_time": "13:45",
			"need_one_time_query": True,
			"one_time_query_time_list": ["08:45", "08:59", "13:25",],
		}
		self.xcfg.update(cfg)

		self.workbook = None
		self.worksheet = None

		self.headers = {'Content-Type': 'application/json'}
		self.url = 'http://%s:%s/option_premium' % (self.xcfg["target_server_address"], self.xcfg["target_server_port"])
		self.saved_cookies = None


	def __enter__(self):
		# Open the workbook
		print "source filepath: %s" % self.xcfg["source_filepath"]
		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		self.worksheet = self.workbook.sheet_by_index(self.xcfg["sheet_index"])
		return self


	def __exit__(self, type, msg, traceback):
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


	def __read_data(self):
		data_dict = {}
		dt_now = datetime.datetime.now()
		print "Read %s at %s" % (os.path.basename(self.xcfg["source_filepath"]), dt_now.strftime(self.DATETIME_FORMAT_STR))
		# import pdb; pdb.set_trace()
		for row_index in range(self.DATA_CELL_ROW_START_INDEX, self.DATA_CELL_ROW_END_INDEX):
			key = None
			try:
				key = self.worksheet.cell_value(row_index, 0)
			except IndexError:
				# print "End row index: %d" % row_index
				break
			# print "row_index: %d, %s" % (row_index, self.worksheet.cell_value(row_index, 0))
			data_dict[key] = {}
			data_cell_column_list = None
			column_index_data = self.DATA_CELL_ROW_COLUMN_INDEX_DICT.get(row_index, None)
			if column_index_data is None:
				data_cell_column_list = range(self.DATA_CELL_COLUMN_DEF_START_INDEX, self.DATA_CELL_COLUMN_DEF_END_INDEX)
			else:
				if type(column_index_data) is dict:
					data_cell_column_list = range(self.DATA_CELL_ROW_COLUMN_INDEX_DICT[row_index]["from"], self.DATA_CELL_ROW_COLUMN_INDEX_DICT[row_index]["to"])
				elif type(column_index_data) is list:
					data_cell_column_list = self.DATA_CELL_ROW_COLUMN_INDEX_DICT[row_index]
				else:
					raise RuntimeError("Unknown column index range in row: %d" % row_index)
			# import pdb; pdb.set_trace()
			for column_index in data_cell_column_list:
				# print "row: %d, column: %d" % (row_index, column_index)
				cell_value = self.worksheet.cell_value(row_index, column_index)
				# print "value: %s" % str(cell_value)
# Check if this option is traded
				if self.__is_string(cell_value):
					data_dict[key][self.DATA_CELL_COLUMN_TITLE_NAME_LIST[column_index]] = None
				# print "%s %d %d" % (key, row_index, column_index)
				# import pdb; pdb.set_trace()
				elif self.DATA_CELL_COLUMN_TYPE_LIST[column_index] == "float":
					data_dict[key][self.DATA_CELL_COLUMN_TITLE_NAME_LIST[column_index]] = float(cell_value)
				elif self.DATA_CELL_COLUMN_TYPE_LIST[column_index] == "int":
					data_dict[key][self.DATA_CELL_COLUMN_TITLE_NAME_LIST[column_index]] = int(cell_value)
				else:
					raise ValueError("Unknown type: %s" % self.DATA_CELL_COLUMN_TYPE_LIST[index])
		# import pdb; pdb.set_trace()
		data_dict["created_at"] = dt_now
		return data_dict


	def __update_data(self, data_dict):
		#print json.dumps(data_dict)	
		res = requests.post(self.url, headers=self.headers, verify=False, cookies=self.saved_cookies, data=json.dumps(data_dict, cls=DayTradeMonitorCollector.DateEncoder))
		#print res.status_code
		if res.status_code in self.HTTP_SUCCESS_CODE_LIST:
			# if len(res.text) > 0: logRead(logData={'log':res.text,'level':'DBG'})
			res_json = json.loads(res.text)
			print "Server[%s] Update: %s" % (self.xcfg["target_server_address"], res_json["msg"])
		else:
			print "Error: %d, %s" % (res.status_code, res.text)
			# sys.exit()


	def collect(self):
		# import pdb; pdb.set_trace()
		data_dict = self.__read_data()
		self.__update_data(data_dict)


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
	parser.add_argument('-f', '--source_filepath', required=False, help='The filepath of data source')
	parser.add_argument('-i', '--sheet_index', required=False, help='The sheet index of data source')
	parser.add_argument('-a', '--target_server_address', required=False, help='The target server address where the date are sent')
	parser.add_argument('-p', '--target_server_port', required=False, help='The target server port where the date are sent')
# How to add option without any argument? use action='store_true'
	'''
	'store_true' and 'store_false' - 这些是 'store_const' 分别用作存储 True 和 False 值的特殊用例。另外，它们的默认值分别为 False 和 True。例如:

	>>> parser = argparse.ArgumentParser()
	>>> parser.add_argument('--foo', action='store_true')
	>>> parser.add_argument('--bar', action='store_false')
	>>> parser.add_argument('--baz', action='store_false')
    '''
	# parser.add_argument('-d', '--disable_check_time', required=False, action='store_true', help='No need to check time for collecting data')
	parser.add_argument('-o', '--one_shot_query', required=False, action='store_true', help='Collect data immediately')
	parser.add_argument('-s', '--start_time', required=False, help='The start time of collecting data. Format: HH:mm')
	parser.add_argument('-e', '--end_time', required=False, help='The end_time of collecting data. Format: HH:mm')
	args = parser.parse_args()
	cfg = {}
	if args.source_filepath is not None:
		cfg['source_filepath'] = args.source_filepath
	if args.sheet_index is not None:
		cfg['sheet_index'] = args.sheet_index
	if args.target_server_address is not None:
		cfg['target_server_address'] = args.target_server_address
	if args.target_server_port is not None:
		cfg['target_server_port'] = args.target_server_port

	# import pdb; pdb.set_trace()
	if args.one_shot_query:
		with DayTradeMonitorCollector(cfg) as obj:
			print "* Collect one shot data at the time [%s]" % datetime.datetime.now()
			obj.collect()
		sys.exit(0)

	start_time_str = args.start_time if (args.start_time is not None) else DEF_START_TIME_STR
	end_time_str = args.end_time if (args.end_time is not None) else DEF_END_TIME_STR
	for _ in CollectorTimer.wait_for_collecting(start_time=start_time_str, end_time=end_time_str):
		with DayTradeMonitorCollector(cfg) as obj:
			obj.collect()
