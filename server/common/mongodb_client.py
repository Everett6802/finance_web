import datetime
import copy
from pymongo import MongoClient
# import common_definition as CMN_DEF
# import common_function as CMN_FUNC
from common import common_definition as CMN_DEF
from common import common_function as CMN_FUNC


class MongoDBClient(object):

    @classmethod
    def __serialize(cls, data):
        serialized_data = copy.deepcopy(data)
        serialized_data["created_at"] = data["created_at"].strftime(CMN_DEF.DB_DATETIME_STRING_FORMAT)
        del serialized_data["_id"]
        return serialized_data


    def __init__(self, collection_name, request_obj, host="localhost", database_name=CMN_DEF.DATABASE_NAME):
        # import pdb; pdb.set_trace()
        self.collection_name = collection_name
        self.request_obj = request_obj
        self.host = host
        self.database_name = database_name
        self.db_client = None
        self.handle = None


    def __enter__(self):
        # import pdb; pdb.set_trace()
        self.db_client = MongoClient('mongodb://%s:27017' % self.host)
        # db = self.db_client.FinanceWebDatabase
        db = self.db_client[self.database_name]
        # self.bid_ask_volume = db["BidAskVolume"]
        self.handle = db[self.collection_name]
        # print "handle connection: %s.%s" % (self.database_name, self.collection_name)    
        return self


    def __exit__(self, type, msg, traceback):
        if self.db_client is not None:
            self.db_client.close()
            self.db_client = None
        return False


    def get_handle(self):
        return self.handle


    def find(self):
        assert self.handle is not None, "self.handle should NOT be None"
        # args = self.request_obj.args
        # print (args) # For debugging
        time_from_criteria = None
        if "from" in self.request_obj.args:
            from_timestr = self.request_obj.args['from']
            from_time_obj = CMN_FUNC.get_datetime_obj_from_query_time_string(from_timestr)
            if from_time_obj is None:
                err_ret = {
                    "msg": "Incorrect start time format: %s" % from_timestr
                }
                return err_ret, 400
            time_from_criteria = {"created_at": {"$gte": from_time_obj}}
        time_until_criteria = None
        if "until" in self.request_obj.args:
            until_timestr = self.request_obj.args['until']
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
        # data_list = self.bid_ask_volume.find(time_criteria).sort("created_at")
        data_list = self.handle.find(time_criteria).sort("created_at")

        # import pdb; pdb.set_trace()
        serialized_data_list = []
        for data in data_list:
            # print self.serialize(data)
            serialized_data_list.append(self.__serialize(data))
        # print "serialized_data_list: %s" % serialized_data_list
        # return jsonify(serialized_data_list)
        return serialized_data_list, 200


    def insert(self):
        assert self.handle is not None, "self.handle should NOT be None"
        # import pdb; pdb.set_trace()
        data = self.request_obj.get_json()
        data["created_at"] = datetime.datetime.strptime(str(data["created_at"]), CMN_DEF.DB_DATETIME_STRING_FORMAT)
        # import pdb; pdb.set_trace()
        # print "data: %s" % data
        # self.bid_ask_volume.insert(data)
        self.handle.insert(data)

        ret = {
            "msg": "New data created at %s" % data["created_at"].strftime(CMN_DEF.DB_DATETIME_STRING_FORMAT)
        }
        return ret, 200


    def delete_many(self):
        assert self.handle is not None, "self.handle should NOT be None"
        # args = self.request_obj.args
        # print (args) # For debugging
        time_from_criteria = None
        if "from" in self.request_obj.args:
            from_timestr = self.request_obj.args['from']
            from_time_obj = CMN_FUNC.get_datetime_obj_from_query_time_string(from_timestr)
            if from_time_obj is None:
                err_ret = {
                    "msg": "Incorrect start time format: %s" % from_timestr
                }
                return err_ret, 400
            time_from_criteria = {"created_at": {"$gte": from_time_obj}}
        time_until_criteria = None
        if "until" in self.request_obj.args:
            until_timestr = self.request_obj.args['until']
            until_time_obj = CMN_FUNC.get_datetime_obj_from_query_time_string(until_timestr)
            if until_time_obj is None:
                err_ret = {
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
        result = self.handle.delete_many(time_criteria)

        # print ("delete count:", result.deleted_count)
        ret = {
            "msg": "%d Data deleted" % result.deleted_count
        }
        return ret, 200
