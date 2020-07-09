#! /usr/bin/python
# -*- coding: utf8 -*-

from flask import jsonify, request
from flask_restful import Resource
import datetime
from common import common_definition as CMN_DEF
from common import common_function as CMN_FUNC
from common import common_class as CMN_CLS


class BidAskVolume(Resource):

	DATETIME_FORMAT_STR = "%Y-%m-%d %H:%M:%S"

	@classmethod
	def serialize(cls, data):
		serialized_data = {
			'created_at': data["created_at"].strftime(cls.DATETIME_FORMAT_STR), 
			"tx_1": data["tx_1"], # 台指期:累計委買(全)
			"tx_2": data["tx_2"], # 台指期:累計委賣(全)
			"tx_3": data["tx_3"], # 台指期:累委買筆(全)
			"tx_4": data["tx_4"], # 台指期:累委賣筆(全)
			"te_1": data["te_1"], # 電子期​:累計委買(全)
			"te_2": data["te_2"], # 電子期​:累計委賣(全)
			"te_3": data["te_3"], # 電子期​:累委買筆(全)
			"te_4": data["te_4"], # 電子期​:累委賣筆(全)
			"tf_1": data["tf_1"], # 金融期:累計委買(全)
			"tf_2": data["tf_2"], # 金融期:累計委賣(全)
			"tf_3": data["tf_3"], # 金融期:累委買筆(全)
			"tf_4": data["tf_4"], # 金融期:累委賣筆(全)
			"cdf_1": data["cdf_1"], # 台積電期:累計委買(全)
			"cdf_2": data["cdf_2"], # 台積電期:累計委賣(全)
			"cdf_3": data["cdf_3"], # 台積電期:累委買筆(全)
			"cdf_4": data["cdf_4"], # 台積電期:累委賣筆(全)
			"dhf_1": data["dhf_1"], # 鴻海期:累計委買(全)
			"dhf_2": data["dhf_2"], # 鴻海期:累計委賣(全)
			"dhf_3": data["dhf_3"], # 鴻海期:累委買筆(全)
			"dhf_4": data["dhf_4"], # 鴻海期:累委賣筆(全)
			"tse_1": data["tse_1"], # 加權:累計委買(全)
			"tse_2": data["tse_2"], # 加權:累計委賣(全)
			"tse_3": data["tse_3"], # 加權:累委買筆(全)
			"tse_4": data["tse_4"], # 加權:累委賣筆(全)
			"tse_5": data["tse_5"], # 加權:上漲家(全)
			"tse_6": data["tse_6"], # 加權:下跌家(全)
		}
		return serialized_data


	def __init__(self):
		# self.db_client = MongoClient('mongodb://localhost:27017')
		# # db = self.db_client.FinanceWebDatabase
		# db = self.db_client[CMN_DEF.DATABASE_NAME]
		# # self.bid_ask_volume = db["BidAskVolume"]
		# self.bid_ask_volume = db["BidAskVolume"]
		pass
		'''
		# How to search database collections in the console

		> use FinanceWebDatabase
		> db.BidAskVolume.find({})
		> db.BidAskVolume.remove({})
		'''


	def test(self):
		start = datetime.datetime(2020, 6, 23, 12, 35, 6, 764)
		end = datetime.datetime(2020, 6, 23, 12, 55, 3, 381)
		# data_list = self.bid_ask_volume.find({'created_at': {'$gte': start, '$lt': end}})
		with CMN_CLS.MongoDBClient(self.__class__.__name__) as db_client:
				data_list = db_client.get_handle().find({'created_at': {'$gte': start, '$lt': end}})

		for data in data_list:
			print self.serialize(data)


	def get(self):
		# assert self.bid_ask_volume is not None, "self.bid_ask_volume should NOT be None"
		# args = request.args
		# print (args) # For debugging
		time_from_criteria = None
		if "from" in request.args:
			from_timestr = request.args['from']
			time_from_criteria = {"created_at": {"$gte": CMN_FUNC.get_datetime_obj_fromm_query_time_strig(from_timestr)}}
		time_until_criteria = None
		if "until" in request.args:
			until_timestr = request.args['until']
			time_until_criteria = {"created_at": {"$lte": CMN_FUNC.get_datetime_obj_fromm_query_time_strig(until_timestr)}}
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
		# data_list = self.bid_ask_volume.find(time_criteria).sort("created_at")
		data_list = None
		with CMN_CLS.MongoDBClient(self.__class__.__name__) as db_client:
			data_list = db_client.get_handle().find(time_criteria).sort("created_at")

		serialized_data_list = []
		for data in data_list:
			print self.serialize(data)
			serialized_data_list.append(self.serialize(data))
			# print "add, new_data_list: %s, len: %d" % (new_data_list, len(new_data_list))
		# print "serialized_data_list: %s" % serialized_data_list
		# return jsonify([self.serialize(data) for data in data_list])
		return jsonify(serialized_data_list)


	def post(self):
		# import pdb; pdb.set_trace()
		# assert self.bid_ask_volume is not None, "self.bid_ask_volume should NOT be None"

		body = request.get_json()
		# created_at = datetime.datetime.now() # datetime.datetime.utcnow()
		data = {
			"created_at": datetime.datetime.strptime(str(body["created_at"]), self.DATETIME_FORMAT_STR),  # created_at,
			"tx_1": body["tx_1"], # 台指期:累計委買(全)
			"tx_2": body["tx_2"], # 台指期:累計委賣(全)
			"tx_3": body["tx_3"], # 台指期:累委買筆(全)
			"tx_4": body["tx_4"], # 台指期:累委賣筆(全)
			"te_1": body["te_1"], # 電子期​:累計委買(全)
			"te_2": body["te_2"], # 電子期​:累計委賣(全)
			"te_3": body["te_3"], # 電子期​:累委買筆(全)
			"te_4": body["te_4"], # 電子期​:累委賣筆(全)
			"tf_1": body["tf_1"], # 金融期:累計委買(全)
			"tf_2": body["tf_2"], # 金融期:累計委賣(全)
			"tf_3": body["tf_3"], # 金融期:累委買筆(全)
			"tf_4": body["tf_4"], # 金融期:累委賣筆(全)
			"cdf_1": body["cdf_1"], # 台積電期:累計委買(全)
			"cdf_2": body["cdf_2"], # 台積電期:累計委賣(全)
			"cdf_3": body["cdf_3"], # 台積電期:累委買筆(全)
			"cdf_4": body["cdf_4"], # 台積電期:累委賣筆(全)
			"dhf_1": body["dhf_1"], # 鴻海期:累計委買(全)
			"dhf_2": body["dhf_2"], # 鴻海期:累計委賣(全)
			"dhf_3": body["dhf_3"], # 鴻海期:累委買筆(全)
			"dhf_4": body["dhf_4"], # 鴻海期:累委賣筆(全)
			"tse_1": body["tse_1"], # 加權:累計委買(全)
			"tse_2": body["tse_2"], # 加權:累計委賣(全)
			"tse_3": body["tse_3"], # 加權:累委買筆(全)
			"tse_4": body["tse_4"], # 加權:累委賣筆(全)
			"tse_5": body["tse_5"], # 加權:上漲家(全)
			"tse_6": body["tse_6"], # 加權:下跌家(全)
		}
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


# ex: http://localhost:5998/bid_ask_volume?from=12:34_20200623&&until=12:56_20200623
	def delete(self):
		# assert self.bid_ask_volume is not None, "self.bid_ask_volume should NOT be None"
		# args = request.args
		# print (args) # For debugging
		time_from_criteria = None
		if "from" in request.args:
			from_timestr = request.args['from']
			time_from_criteria = {"created_at": {"$gte": CMN_FUNC.get_datetime_obj_fromm_query_time_strig(from_timestr)}}
		time_until_criteria = None
		if "until" in request.args:
			until_timestr = request.args['until']
			time_until_criteria = {"created_at": {"$lte": CMN_FUNC.get_datetime_obj_fromm_query_time_strig(until_timestr)}}
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
