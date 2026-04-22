#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import argparse
import requests
import pandas as pd
import numpy as np
import time
from datetime import datetime, date, timedelta


class AlphaStackAnalyzer(object):

	DEFAULT_FINMIND_DATE_FORMAT = "%Y-%m-%d"
	DEFAULT_FINMIND_MIN_DATE = date(1900, 1, 1)
	DEFAULT_FINMIND_MIN_DATE_STR = DEFAULT_FINMIND_MIN_DATE.strftime(DEFAULT_FINMIND_DATE_FORMAT)

# ========================
# 基礎抓資料函數
# ========================
	@classmethod
	def __get_data(cls, token, dataset_name, stock_symbol, fetch_start=None, fetch_end=None):
		url = "https://api.finmindtrade.com/api/v4/data"
		params = {
			"dataset": dataset_name,
			"token": token,
			"data_id": stock_symbol,
		}
		if fetch_start is not None:
			params["start_date"] = fetch_start.strftime(cls.DEFAULT_FINMIND_DATE_FORMAT)
		else:
			params["start_date"] = cls.DEFAULT_FINMIND_MIN_DATE_STR
		if fetch_end is not None:
			params["end_date"] = fetch_end.strftime(cls.DEFAULT_FINMIND_DATE_FORMAT)

		data = requests.get(url, params=params).json()
		return pd.DataFrame(data["data"])


# ========================
# ROE & FCF（財報）
# ========================
	@classmethod
	def __get_financial_factors(cls, token, stock_symbol, fetch_start=None, fetch_end=None):
		df = cls.__get_data(token, "TaiwanStockFinancialStatements", stock_symbol, fetch_start, fetch_end)
		df = df.pivot_table(index="date",columns="type",values="value",aggfunc="first").reset_index()
		df["ROE"] = df["NetIncome"] / df["Equity"]
		df["FCF"] = df["CashFlowsFromOperatingActivities"] - df["CapitalExpenditures"]
		return df[["date", "ROE", "FCF"]]


# ========================
# PEGY（TTM）
# ========================
	@classmethod
	def __get_pegy(cls, token, stock_symbol, fetch_start=None, fetch_end=None):
# PE
		pe = cls.__get_data(token, "TaiwanStockPER", stock_symbol, fetch_start, fetch_end)
		pe = pe.rename(columns={"PER": "pe"})
		pe["date"] = pd.to_datetime(pe["date"])
# 月營收
		revenue = cls.__get_data(token, "TaiwanStockMonthRevenue", stock_symbol, fetch_start, fetch_end)
		revenue["date"] = pd.to_datetime(revenue["date"])
		revenue = revenue.sort_values("date")
# TTM營收
		revenue["TTM_revenue"] = revenue["revenue"].rolling(12).sum()
		revenue["TTM_YoY"] = revenue["TTM_revenue"].pct_change(12)
# 股利
		dividend = cls.__get_data(token, "TaiwanStockDividend", stock_symbol, fetch_start, fetch_end)
		dividend["date"] = pd.to_datetime(dividend["date"])
		dividend["cash_dividend"] = dividend["CashDividend"]
# forward fill
		dividend = dividend.sort_values("date")
		dividend["TTM_dividend"] = dividend["cash_dividend"]
		dividend["TTM_dividend"] = dividend["TTM_dividend"].fillna(method="ffill")
# 股價
		price = cls.__get_data(token, "TaiwanStockPrice", stock_symbol, fetch_start, fetch_end)
		price["date"] = pd.to_datetime(price["date"])
# 合併（月頻）
		df = pe.merge(revenue, on="date", how="left")
		df = df.merge(dividend[["date", "TTM_dividend"]], on="date", how="left")
		df = df.merge(price[["date", "close"]], on="date", how="left")
		df = df.sort_values("date")
# forward fill
		df = df.fillna(method="ffill")
# 殖利率
		df["yield"] = df["TTM_dividend"] / df["close"]
# PEGY
		df["PEGY"] = df["pe"] / (df["TTM_YoY"] + df["yield"])
		return df[["date", "PEGY"]]


# ========================
# Sharpe Ratio
# ========================
	@classmethod
	def __get_sharpe(cls, token, stock_symbol, fetch_start=None, fetch_end=None):
		price = cls.__get_data(token, "TaiwanStockPrice", stock_symbol, fetch_start, fetch_end)
		price["date"] = pd.to_datetime(price["date"])
		price = price.sort_values("date")
		price["return"] = price["close"].pct_change()
		window = 252
		price["Sharpe"] = (price["return"].rolling(window).mean() / price["return"].rolling(window).std()) * np.sqrt(252)
		return price[["date", "Sharpe"]]


	def __init__(self, cfg):
		self.xcfg = {
			"finmind_token": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)


	def __enter__(self):
		return self


	def __exit__(self, type, msg, traceback):
		return False


	def analyze(self, stock_symbol, fetch_start=None, fetch_end=None):
		df1 = self.__get_financial_factors(self.xcfg["finmind_token"], stock_symbol, fetch_start, fetch_end)
		df2 = self.__get_pegy(self.xcfg["finmind_token"], stock_symbol, fetch_start, fetch_end)
		df3 = self.__get_sharpe(self.xcfg["finmind_token"], stock_symbol, fetch_start, fetch_end)
# merge
		df = df2.merge(df1, on="date", how="left")
		df = df.merge(df3, on="date", how="left")
		df = df.sort_values("date")
# 避免未來函數
		df["Sharpe"] = df["Sharpe"].shift(1)
# 清理資料
		df = df.replace([np.inf, -np.inf], np.nan)
		df = df.dropna()
# 篩選條件（可調整）
		df_filtered = df[
		    (df["PEGY"] < 1) &
		    (df["ROE"] > 0.15) &
		    (df["FCF"] > 0) &
		    (df["Sharpe"] > 1)
		]
		print(df_filtered.tail())
		df.to_csv("factor_data.csv", index=False)


if __name__ == "__main__":
# argparse 預設會把 help 文字裡的換行與多重空白「壓縮」成一行，所以你在字串裡寫的 \n 不一定會照原樣顯示。 => 建立 parser 時加上 formatter_class=argparse.RawTextHelpFormatter
	parser = argparse.ArgumentParser(description='Print help', formatter_class=argparse.RawTextHelpFormatter)
	'''
	參數基本上分兩種，一種是位置參數 (positional argument)，另一種就是選擇性參數 (optional argument)
	* example2.py
	parser.add_argument("pos1", help="positional argument 1")
	parser.add_argument("-o", "--optional-arg", help="optional argument", dest="opt", default="default")

	# python example2.py hello -o world 
	positional arg: hello
	optional arg: world
	'''
# How to add option without any argument? use action='store_true'
	'''
	'store_true' and 'store_false' - ?些是 'store_const' 分?用作存? True 和 False 值的特殊用例。
	另外，它?的默?值分?? False 和 True。例如:

	>>> parser = argparse.ArgumentParser()
	>>> parser.add_argument('--foo', action='store_true')
	>>> parser.add_argument('--bar', action='store_false')
	>>> parser.add_argument('--baz', action='store_false')
	'''
	parser.add_argument('--finmind_token', required=False, help='The FinMind Token')
	args = parser.parse_args()
	# import pdb; pdb.set_trace()
	cfg = {}
	if args.finmind_token is not None: cfg['finmind_token'] = args.finmind_token
	# import pdb; pdb.set_trace()
	with AlphaStackAnalyzer(cfg) as obj:
		if args.analyze:
			obj.analyze()
			sys.exit(0)