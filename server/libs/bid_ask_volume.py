#! /usr/bin/python
# -*- coding: utf8 -*-

from flask import jsonify, request
from flask_restful import Resource
# import datetime
import copy
from common import common_definition as CMN_DEF
from common import common_function as CMN_FUNC
from common.common_variable import GlobalVar as GV
from common import mongodb_client as MONGODB


class BidAskVolume(Resource):

	# DATETIME_FORMAT_STR = "%Y-%m-%d %H:%M:%S"

	# @classmethod
	# def serialize(cls, data):
	# 	# serialized_data = {
	# 	# 	'created_at': data["created_at"].strftime(cls.DATETIME_FORMAT_STR), 
	# 	# 	u"台指期:累計委買(全)": data[u"台指期:累計委買(全)"], # 台指期:累計委買(全)
	# 	# 	u"台指期:累計委賣(全)": data[u"台指期:累計委賣(全)"], # 台指期:累計委賣(全)
	# 	# 	# "tx_1": data["tx_1"], # 台指期:累計委買(全)
	# 	# 	# "tx_2": data["tx_2"], # 台指期:累計委賣(全)
	# 	# 	"tx_3": data["tx_3"], # 台指期:累委買筆(全)
	# 	# 	"tx_4": data["tx_4"], # 台指期:累委賣筆(全)
	# 	# 	"te_1": data["te_1"], # 電子期​:累計委買(全)
	# 	# 	"te_2": data["te_2"], # 電子期​:累計委賣(全)
	# 	# 	"te_3": data["te_3"], # 電子期​:累委買筆(全)
	# 	# 	"te_4": data["te_4"], # 電子期​:累委賣筆(全)
	# 	# 	"tf_1": data["tf_1"], # 金融期:累計委買(全)
	# 	# 	"tf_2": data["tf_2"], # 金融期:累計委賣(全)
	# 	# 	"tf_3": data["tf_3"], # 金融期:累委買筆(全)
	# 	# 	"tf_4": data["tf_4"], # 金融期:累委賣筆(全)
	# 	# 	"cdf_1": data["cdf_1"], # 台積電期:累計委買(全)
	# 	# 	"cdf_2": data["cdf_2"], # 台積電期:累計委賣(全)
	# 	# 	"cdf_3": data["cdf_3"], # 台積電期:累委買筆(全)
	# 	# 	"cdf_4": data["cdf_4"], # 台積電期:累委賣筆(全)
	# 	# 	"dhf_1": data["dhf_1"], # 鴻海期:累計委買(全)
	# 	# 	"dhf_2": data["dhf_2"], # 鴻海期:累計委賣(全)
	# 	# 	"dhf_3": data["dhf_3"], # 鴻海期:累委買筆(全)
	# 	# 	"dhf_4": data["dhf_4"], # 鴻海期:累委賣筆(全)
	# 	# 	"tse_1": data["tse_1"], # 加權:累計委買(全)
	# 	# 	"tse_2": data["tse_2"], # 加權:累計委賣(全)
	# 	# 	"tse_3": data["tse_3"], # 加權:累委買筆(全)
	# 	# 	"tse_4": data["tse_4"], # 加權:累委賣筆(全)
	# 	# 	"tse_5": data["tse_5"], # 加權:上漲家(全)
	# 	# 	"tse_6": data["tse_6"], # 加權:下跌家(全)
	# 	# }
	# 	serialized_data = copy.deepcopy(data)
	# 	serialized_data["created_at"] = data["created_at"].strftime(cls.DATETIME_FORMAT_STR)
	# 	del serialized_data["_id"]
	# 	return serialized_data


	def __init__(self):
		# self.db_client = MongoClient('mongodb://localhost:27017')
		# # db = self.db_client.FinanceWebDatabase
		# db = self.db_client[CMN_DEF.DATABASE_NAME]
		# self.bid_ask_volume = db["BidAskVolume"]
		pass
		'''
		# How to search database collections in the console

		> use FinanceWebDatabase
		> db.BidAskVolume.find({})
		> db.BidAskVolume.remove({})
		'''


	def get(self):
		ret_data = None
		ret_http_code = None
		with MONGODB.MongoDBClient(self.__class__.__name__, request, host=GV.HOSTNAME) as db_client:
			ret_data, ret_http_code = db_client.find()
		return ret_data, ret_http_code



	def post(self):
		ret_data = None
		ret_http_code = None
		with MONGODB.MongoDBClient(self.__class__.__name__, request, host=GV.HOSTNAME) as db_client:
			ret_data, ret_http_code = db_client.insert()
		return ret_data, ret_http_code


# ex: http://localhost:5998/bid_ask_volume?from=12:34_20200623&&until=12:56_20200623
	def delete(self):
		ret_data = None
		ret_http_code = None
		with MONGODB.MongoDBClient(self.__class__.__name__, request, host=GV.HOSTNAME) as db_client:
			ret_data, ret_http_code = db_client.delete_many()
		return ret_data, ret_http_code
