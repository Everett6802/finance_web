#! /usr/bin/python
# -*- coding: utf8 -*-

import os

import xlrd
# import xlsxwriter
import argparse


class TakeProfitTracker(object):

	DEFAULT_SOURCE_FOLDERPATH =  "C:\\停利追蹤"
	DEFAULT_SOURCE_FILENAME = "take_profile_tracker"
	DEFAULT_SOURCE_FULL_FILENAME = "%s.xlsx" % DEFAULT_SOURCE_FILENAME

	def __init__(self, cfg):
		self.xcfg = {
			"source_folderpath": None,
			"source_filename": self.DEFAULT_SOURCE_FULL_FILENAME,
		}
		self.xcfg.update(cfg)
		self.xcfg["source_folderpath"] = self.DEFAULT_SOURCE_FOLDERPATH if self.xcfg["source_folderpath"] is None else self.xcfg["source_folderpath"]
		self.xcfg["source_filename"] = self.DEFAULT_SOURCE_FULL_FILENAME if self.xcfg["source_filename"] is None else self.xcfg["source_filename"]
		self.xcfg["source_filepath"] = os.path.join(self.xcfg["source_folderpath"], self.xcfg["source_filename"])


	def __enter__(self):
# Open the workbook
		self.workbook = xlrd.open_workbook(self.xcfg["source_filepath"])
		return self


	def __exit__(self, type, msg, traceback):
		if self.workbook is not None:
			self.workbook.release_resources()
			del self.workbook
			self.workbook = None
		return False


if __name__ == "__main__":
	parser = argparse.ArgumentParser(description='Print help')
	
	cfg = {}
	
	with TakeProfitTracker(cfg) as obj:
		pass
