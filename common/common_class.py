from pymongo import MongoClient
import common_definition as CMN_DEF


class MongoDBClient(object):

    def __init__(self, collection_name, host="localhost", database_name=CMN_DEF.DATABASE_NAME):
        # import pdb; pdb.set_trace()
        self.collection_name = collection_name
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
        return False


    def get_handle(self):
    	return self.handle
