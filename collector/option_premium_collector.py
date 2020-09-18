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

class OptionPremiumCollector(object):

	DEFAULT_SOURCE_FILENAME = "xq_option_premium.xlsx"
	DEFAULT_SHEET_INDEX = 0
	DATA_CELL_CHECK_NUMBER_COLUMN_INDEX = 4
	DATA_CELL_COLUMN_LIST = [4, 7, 8,] # [成交, 總量, 未平倉,]
	DATA_CELL_COLUMN_TYPE_LIST = ["float", "int", "int", ] # [成交, 總量, 未平倉,]
	DATA_CELL_ROW_START_INDEX = 1
	DATA_CELL_ROW_END_MAX_INDEX = 200
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
				return obj.strftime(OptionPremiumCollector.DATETIME_FORMAT_STR)  
			elif isinstance(obj, date):  
				return obj.strftime(OptionPremiumCollector.DATE_FORMAT_STR)  
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
		}
		self.xcfg.update(cfg)

		self.workbook = None
		self.worksheet = None

		self.headers = {'Content-Type': 'application/json'}
		self.url = 'http://%s:%s/option_premium' % (self.xcfg["target_server_address"], self.xcfg["target_server_port"])
		self.saved_cookies = None


	def __enter__(self):
		# Open the workbook
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
		# for key, value in self.DATA_CELL_COORDINATE_DICT.items():
		# 	row_index, column_index = value
		# 	data_dict[key] = int(self.worksheet.cell_value(row_index, column_index))

		for row_index in range(self.DATA_CELL_ROW_START_INDEX, self.DATA_CELL_ROW_END_MAX_INDEX):
			key = None
			try:
				key = self.worksheet.cell_value(row_index, 0)
			except IndexError:
				# print "End row index: %d" % row_index
				break
			# print "row_index: %d, %s" % (row_index, self.worksheet.cell_value(row_index, 0))
			data_dict[key] = []
			for index, column_index in enumerate(self.DATA_CELL_COLUMN_LIST):
# Check if this option is traded
				if self.__is_string(self.worksheet.cell_value(row_index, self.DATA_CELL_CHECK_NUMBER_COLUMN_INDEX)):
					continue
				# print "%s %d %d" % (key, row_index, column_index)
				# import pdb; pdb.set_trace()
				if self.DATA_CELL_COLUMN_TYPE_LIST[index] == "float":
					data_dict[key].append(float(self.worksheet.cell_value(row_index, column_index)))
				elif self.DATA_CELL_COLUMN_TYPE_LIST[index] == "int":
					data_dict[key].append(int(self.worksheet.cell_value(row_index, column_index)))
				else:
					raise ValueError("Unknown type: %s" % self.DATA_CELL_COLUMN_TYPE_LIST[index])

		# import pdb; pdb.set_trace()
		data_dict["created_at"] = dt_now
		return data_dict


	def __update_data(self, data_dict):
		#print json.dumps(data_dict)	
		res = requests.post(self.url, headers=self.headers, verify=False, cookies=self.saved_cookies, data=json.dumps(data_dict, cls=OptionPremiumCollector.DateEncoder))
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
	# import pdb; pdb.set_trace()
	# res = requests.get("http://10.206.24.219:5998/test")

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
		with OptionPremiumCollector(cfg) as obj:
			print "* Collect one shot data at the time [%s]" % datetime.datetime.now()
			obj.collect()
		sys.exit(0)

	start_time_str = args.start_time if (args.start_time is not None) else DEF_START_TIME_STR
	end_time_str = args.end_time if (args.end_time is not None) else DEF_END_TIME_STR
	for _ in CollectorTimer.wait_for_collecting(start_time=start_time_str, end_time=end_time_str):
		with OptionPremiumCollector(cfg) as obj:
			obj.collect()

	# dt_start = None
	# dt_end = None
	# # import pdb; pdb.set_trace()
	# if not args.disable_check_time:
	# 	start_time_str = args.start_time if (args.start_time is not None) else DEF_START_TIME_STR
	# 	end_time_str = args.end_time if (args.end_time is not None) else DEF_END_TIME_STR

	# 	dt_now = datetime.datetime.now()
	# 	dt_start_time_tmp = datetime.datetime.strptime(start_time_str, "%H:%M")
	# 	dt_end_time_tmp = datetime.datetime.strptime(end_time_str, "%H:%M")
	# 	dt_today_start = datetime.datetime(dt_now.year, dt_now.month, dt_now.day, dt_start_time_tmp.hour, dt_start_time_tmp.minute, 0)
	# 	dt_today_end = datetime.datetime(dt_now.year, dt_now.month, dt_now.day, dt_end_time_tmp.hour, dt_end_time_tmp.minute, 0)
	# 	# import pdb; pdb.set_trace()
	# 	if dt_now >= dt_today_end:
	# 		print "Current time[%s] expires on the end time [%s]" % (dt_now, dt_today_end)
	# 		sys.exit(0)
	# 	elif dt_now <= dt_today_start:
	# 		dt_start = dt_today_start
	# 	else:
	# 		time_diff = (int((dt_now - dt_today_start).total_seconds()) / DEF_TIME_INTERVAL_IN_SEC + 1) * DEF_TIME_INTERVAL_IN_SEC
	# 		dt_start = dt_today_start + datetime.timedelta(seconds=time_diff)
	# 	dt_end = dt_today_end
	# 	wait_time_before_start = (dt_start - datetime.datetime.now()).total_seconds()
	# 	print "* Collect data in time range[%s, %s] * " % (dt_start, dt_end)
	# 	print "Wait %d seconds before start......" % wait_time_before_start
	# 	time.sleep(wait_time_before_start)

	# while True:
	# 	if not args.disable_check_time:
	# 		 dt_now = datetime.datetime.now()
	# 		 if dt_now > dt_end:
	# 		 	print "Current time[%s] is NOT in the range [%s, %s]... STOP" % (dt_now, dt_start, dt_end)
	# 		 	break
	# 	with OptionPremiumCollector(cfg) as obj:
	# 		obj.collect()
	# 	time.sleep(DEF_TIME_INTERVAL_IN_SEC)

	# if not args.disable_check_time:
	# 	print "* Collect data in time range[%s, %s]... DONE" % (dt_start, dt_end)
