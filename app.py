from flask import Flask, jsonify, request
from flask_restful import Api, Resource
from libs.test import Test
from libs.bid_ask_volume import BidAskVolume
from libs.stock_futures_bid_ask_volume import StockFuturesBidAskVolume


app = Flask(__name__)
api = Api(app)

api.add_resource(Test, "/test")
api.add_resource(BidAskVolume, "/bid_ask_volume")
api.add_resource(StockFuturesBidAskVolume, "/stock_futures_bid_ask_volume")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5998, debug=False)
