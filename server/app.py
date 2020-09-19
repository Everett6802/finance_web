import argparse
from flask import Flask, jsonify, request
from flask_restful import Api, Resource
from common.common_variable import GlobalVar as GV
from libs.test import Test
from libs.bid_ask_volume import BidAskVolume
from libs.stock_futures_bid_ask_volume import StockFuturesBidAskVolume


app = Flask(__name__)
api = Api(app)

api.add_resource(Test, "/test")
api.add_resource(BidAskVolume, "/bid_ask_volume")
api.add_resource(StockFuturesBidAskVolume, "/stock_futures_bid_ask_volume")
api.add_resource(OptionPremium, "/option_preminum")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Print help')
    parser.add_argument('-n', '--hostname', required=False, help='The hostname of database')
    args = parser.parse_args()
    if args.hostname is not None:
        GV.HOSTNAME = args.hostname

    GV.GLOBAL_VARIABLE_UPDATED = True
    # print GV.HOSTNAME
    app.run(host="0.0.0.0", port=5998, debug=False)
