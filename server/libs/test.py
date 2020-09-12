#! /usr/bin/python
# -*- coding: utf8 -*-

from flask import jsonify, request
from flask_restful import Resource
from pymongo import MongoClient
import datetime
from common import common_function as CMN_FUNC
from common.common_variable import GlobalVar as GV


class Test(Resource):

	@classmethod
	def serialize(cls, data):
		serialized_data = {
			'created_at': data["created_at"].strftime("%Y-%m-%d %H:%M:%S"), 
			'code': data["code"]
		}
		return serialized_data


	def __init__(self):
		self.db_client = MongoClient('mongodb://%s:27017' % GV.HOSTNAME)
		db = self.db_client.TestDatabase
		self.test = db["Test"]


	def get(self):
		return {"test": "Test Only",}


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
			"msg": "New data created at %s" % created_at.strftime("%Y-%m-%d %H:%M:%S")
		}
		return jsonify(ret)
