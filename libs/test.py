#! /usr/bin/python
# -*- coding: utf8 -*-

from flask import jsonify, request
from flask_restful import Resource
from pymongo import MongoClient
import datetime
from common import common_function as CMN_FUNC


class Test(Resource):

	@classmethod
	def serialize(cls, data):
		serialized_data = {
			'created_at': data["created_at"].strftime("%Y-%m-%d %H:%M:%S"), 
			'code': data["code"]
		}
		return serialized_data


	def __init__(self):
		self.db_client = MongoClient('mongodb://localhost:27017')
		db = self.db_client.TestDatabase
		self.test = db["Test"]


	def get(self):
		assert self.test is not None, "self.test should NOT be None"
		# args = request.args
		# print (args) # For debugging
# Set the search time range
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
		data_list = self.test.find(time_criteria).sort("created_at")
		serialized_data_list = []
		for data in data_list:
			serialized_data_list.append(self.serialize(data))
			# print "add, new_data_list: %s, len: %d" % (new_data_list, len(new_data_list))
		# print "serialized_data_list: %s" % serialized_data_list
		# return jsonify([self.serialize(data) for data in data_list])
		return jsonify(serialized_data_list)


	def post(self):
		assert self.test is not None, "self.test should NOT be None"

		body = request.get_json()
		# print body
		created_at = datetime.datetime.now() # datetime.datetime.utcnow()
		data = {
			"created_at": created_at,
			"code": body["code"],
		}
		self.test.insert(data)
		ret = {
			"status": 200,
			"msg": "New data created at %s" % created_at.strftime("%Y-%m-%d %H:%M:%S")
		}
		return jsonify(ret)
