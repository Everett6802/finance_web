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


class OptionPremium(Resource):

	# DATETIME_FORMAT_STR = "%Y-%m-%d %H:%M:%S"


	def __init__(self):
		# self.db_client = MongoClient('mongodb://localhost:27017')
		# # db = self.db_client.FinanceWebDatabase
		# db = self.db_client[CMN_DEF.DATABASE_NAME]
		# self.bid_ask_volume = db["OptionPremium"]
		pass
		'''
		# How to search database collections in the console

		> use FinanceWebDatabase
		> db.OptionPremium.find({})
		> db.OptionPremium.remove({})
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
