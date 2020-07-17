#! /usr/bin/python
# -*- coding: utf8 -*-

from flask import jsonify, request
from flask_restful import Resource
import datetime
import copy
from common import common_definition as CMN_DEF
from common import common_function as CMN_FUNC
from common import common_class as CMN_CLS


class StockFuturesBidAskVolume(Resource):

	DATETIME_FORMAT_STR = "%Y-%m-%d %H:%M:%S"

	@classmethod
	def serialize(cls, data):
		serialized_data = copy.deepcopy(data)
		serialized_data["created_at"] = data["created_at"].strftime(cls.DATETIME_FORMAT_STR)
		del serialized_data["_id"]
		return serialized_data


	def __init__(self):
		# self.db_client = MongoClient('mongodb://localhost:27017')
		# # db = self.db_client.FinanceWebDatabase
		# db = self.db_client[CMN_DEF.DATABASE_NAME]
		# self.stock_futures_bid_ask_volume = db["StockFuturesBidAskVolume"]
		pass
		'''
		# How to search database collections in the console

		> use FinanceWebDatabase
		> db.BidAskVolume.find({})
		> db.BidAskVolume.remove({})
		'''


	def test(self):
		pass


# http://localhost:5998/stock_futures_bid_ask_volume?from=16:50_20200716&&from=17:50_20200716
	def get(self):
		time_from_criteria = None
		if "from" in request.args:
			from_timestr = request.args['from']
			from_time_obj = CMN_FUNC.get_datetime_obj_from_query_time_string(from_timestr)
			if from_time_obj is None:
				err_ret = {
					# "status": 400,
					"msg": "Incorrect start time format: %s" % from_timestr
				}
				return err_ret, 400
			time_from_criteria = {"created_at": {"$gte": from_time_obj}}
		time_until_criteria = None
		if "until" in request.args:
			until_timestr = request.args['until']
			until_time_obj = CMN_FUNC.get_datetime_obj_from_query_time_string(until_timestr)
			if until_time_obj is None:
				err_ret = {
					# "status": 400,
					"msg": "Incorrect end time format: %s" % until_timestr
				}
				return err_ret, 400
			time_until_criteria = {"created_at": {"$lte": until_time_obj}}
		time_criteria = None
		if time_from_criteria is not None and time_until_criteria is not None:
			time_criteria = {"$and": [time_from_criteria, time_until_criteria]}
		elif time_from_criteria is not None:
			time_criteria = time_from_criteria
		elif time_until_criteria is not None:
			time_criteria = time_until_criteria
		else:
			time_criteria = {}
		print time_criteria
		# data_list = self.bid_ask_volume.find(time_criteria).sort("created_at")
		data_list = None
		with CMN_CLS.MongoDBClient(self.__class__.__name__) as db_client:
			data_list = db_client.get_handle().find(time_criteria).sort("created_at")

		# import pdb; pdb.set_trace()
		serialized_data_list = []
		for data in data_list:
			# print self.serialize(data)
			serialized_data_list.append(self.serialize(data))
		# print "serialized_data_list: %s" % serialized_data_list
		return jsonify(serialized_data_list)


	def post(self):
		# import pdb; pdb.set_trace()
		data = request.get_json()
		data["created_at"] = datetime.datetime.strptime(str(data["created_at"]), self.DATETIME_FORMAT_STR)
		# import pdb; pdb.set_trace()
		# print "data: %s" % data
		# self.bid_ask_volume.insert(data)
		with CMN_CLS.MongoDBClient(self.__class__.__name__) as db_client:
			db_client.get_handle().insert(data)

		ret = {
			"status": 200,
			"msg": "New data created at %s" % data["created_at"].strftime(self.DATETIME_FORMAT_STR)
		}
		return jsonify(ret)


# ex: http://localhost:5998/stock_futures_bid_ask_volume?from=12:34_20200623&&until=12:56_20200623
	def delete(self):
		# assert self.bid_ask_volume is not None, "self.bid_ask_volume should NOT be None"
		# args = request.args
		# print (args) # For debugging
		time_from_criteria = None
		if "from" in request.args:
			from_timestr = request.args['from']
			from_time_obj = CMN_FUNC.get_datetime_obj_from_query_time_string(from_timestr)
			if from_time_obj is None:
				err_ret = {
					# "status": 400,
					"msg": "Incorrect start time format: %s" % from_timestr
				}
				return err_ret, 400
			time_from_criteria = {"created_at": {"$gte": from_time_obj}}
		time_until_criteria = None
		if "until" in request.args:
			until_timestr = request.args['until']
			until_time_obj = CMN_FUNC.get_datetime_obj_from_query_time_string(until_timestr)
			if until_time_obj is None:
				err_ret = {
					# "status": 400,
					"msg": "Incorrect end time format: %s" % until_timestr
				}
				return err_ret, 400
			time_until_criteria = {"created_at": {"$lte": until_time_obj}}
		time_criteria = None
		if time_from_criteria is not None and time_until_criteria is not None:
			time_criteria = {"$and": [time_from_criteria, time_until_criteria]}
		elif time_from_criteria is not None:
			time_criteria = time_from_criteria
		elif time_until_criteria is not None:
			time_criteria = time_until_criteria
		else:
			time_criteria = {}
		# print time_criteria
		# import pdb; pdb.set_trace()
# The Remove() method is out of date, 
# so use delete_one() or delete_many() to delete documents
		# write_res = self.bid_ask_volume.remove(time_criteria)
		# result = self.bid_ask_volume.delete_many(time_criteria)
		result = None
		with CMN_CLS.MongoDBClient(self.__class__.__name__) as db_client:
			result = db_client.get_handle().delete_many(time_criteria)

		# print ("delete count:", result.deleted_count)
		ret = {
			"status": 200,
			"msg": "%d Data deleted" % result.deleted_count
		}
		return jsonify(ret)
