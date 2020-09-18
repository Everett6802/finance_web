#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import time
import datetime
from lib import cmn_definition as CMN_DEF
# import cmn_definition as CMN_DEF


DEF_START_TIME_STR = "08:30"
DEF_END_TIME_STR = "13:46"
DEF_TIME_INTERVAL_IN_SEC = 300

class CollectorTimer(object):

	@classmethod
	def _get_today_datetime_from_time_str(cls, time_str, dt_now=None, time_str_format=CMN_DEF.TIME_FORAMT_STR):
		if dt_now is None:
			dt_now = datetime.datetime.now()
		dt_time_tmp = datetime.datetime.strptime(time_str, time_str_format)
		dt_today_time = datetime.datetime(dt_now.year, dt_now.month, dt_now.day, dt_time_tmp.hour, dt_time_tmp.minute, 0)
		return dt_today_time


	@classmethod
	def wait_for_collecting(cls, **kwargs):
		collector_timer = cls(kwargs)
		collector_timer._generate_time_slice()
		dt_now = datetime.datetime.now()
# Filter the expired time slice
		collector_timer.time_slice_list = filter(lambda x: x > dt_now, collector_timer.time_slice_list)
		collector_timer.time_slice_list_len = len(collector_timer.time_slice_list)
		if collector_timer.time_slice_list_len == 0:
			print "* Time slice list is EMPTY *\n"
			return
		else:
			print "* Collect data in time range[%s, %s] *\n" % (collector_timer.StartTime, collector_timer.EndTime)

		for index, dt_time_slice in enumerate(collector_timer):
			dt_now = datetime.datetime.now()
			sleep_time = (dt_time_slice - dt_now).total_seconds()
			# print "Current time: %s, Next collection time: %s, time diff: %d" % (dt_now, dt_time_slice, sleep_time)
			if index == 0:
				print "Wait %d seconds before start......" % sleep_time
			# else:
			# 	print "Wait %d seconds for the next iteration" % sleep_time
			time.sleep(sleep_time)
			yield
		print "Current time[%s] is NOT in the range [%s, %s]... STOP" % (datetime.datetime.now(), collector_timer.StartTime, collector_timer.EndTime)


	def __init__(self, cfg):
		self.xcfg = {
			"start_time": DEF_START_TIME_STR,
			"end_time": DEF_END_TIME_STR,
			"need_periodic_query": True,
			"periodic_query_time_interval": DEF_TIME_INTERVAL_IN_SEC,  # Unit: second
			"need_one_time_query": False,
			"one_time_query_time_list": [],
			"remove_duplicate_time": True,
		}
		self.xcfg.update(cfg)

		self.time_slice_list = None
		self.time_slice_list_index = 0
		self.time_slice_list_len = 0

		self._generate_time_slice()



	def __str__(self):
		self_str = ""
		for index, dt_time_slice in enumerate(self.time_slice_list):
			self_str += "%03d => %s\n" % (index, time_slice)
		return self_str


	def __iter__(self):
		return self


	def next(self):
		if self.time_slice_list_index == self.time_slice_list_len:
# Finish the iteration and reset the index
			self.time_slice_list_index = 0
			raise StopIteration
		time_slice = self.time_slice_list[self.time_slice_list_index]
		self.time_slice_list_index += 1
		return time_slice


	def _generate_time_slice(self):
		self.time_slice_list = []

		dt_now = datetime.datetime.now()
		dt_today_start = self._get_today_datetime_from_time_str(self.xcfg["start_time"], dt_now)
		dt_today_end = self._get_today_datetime_from_time_str(self.xcfg["end_time"], dt_now)
		if dt_today_start > dt_today_end:
			print "WARNING: The start time[%s] is later than the end time[%s]" % (dt_today_start, dt_today_end)

		if self.xcfg["need_periodic_query"]:
			self.time_slice_list.append(dt_today_start)
			cnt = 1
			while True:
				dt_current = dt_today_start + datetime.timedelta(seconds = self.xcfg["periodic_query_time_interval"] * cnt)
				if dt_current > dt_today_end:
					break
				self.time_slice_list.append(dt_current)
				cnt += 1
		if self.xcfg["need_one_time_query"]:
			for query_time in self.xcfg["one_time_query_time_list"]:
				# dt_one_time_tmp = datetime.datetime.strptime(query_time, "%H:%M")
				# dt_one_time = datetime.datetime(dt_now.year, dt_now.month, dt_now.day, dt_one_time_tmp.hour, dt_one_time_tmp.minute, 0)
				dt_one_time = self._get_today_datetime_from_time_str(query_time, dt_now)
				self.time_slice_list.append(dt_one_time)
# Remove the duplicate time
			if self.xcfg["remove_duplicate_time"]:
				self.time_slice_list = list(set(self.time_slice_list))
			self.time_slice_list.sort()

		self.time_slice_list_len = len(self.time_slice_list)


	@property
	def StartTime(self):
		return self.time_slice_list[0]


	@property
	def EndTime(self):
		return self.time_slice_list[-1]


if __name__ == "__main__":
	cfg = {
		"need_one_time_query": True,
		"one_time_query_time_list": ["11:10","08:47","13:24","10:59",],
	}
	# while next(CollectorTimer.wait_for_collecting(**cfg)):
	# 	pass
	for _ in CollectorTimer.wait_for_collecting(**cfg):
		pass

	# my_iterator = MyIterator(3)

	# for _ in my_iterator:
	# 	pass
		# print(item)

	# collector_timer = CollectorTimer(cfg)
	# collector_timer.initialize()

	# # print (collector_timer)

	# dt_now = datetime.datetime.now()

	# print "Check1"
	# for index, dt_time_slice in enumerate(collector_timer):
	# 	print "%03d => %s, %d" % (index, dt_time_slice, (dt_now - dt_time_slice).total_seconds())

	# print "Check2"
	# for index, dt_time_slice in enumerate(collector_timer):
	# 	print "%03d => %s, %d" % (index, dt_time_slice, (dt_now - dt_time_slice).total_seconds())
