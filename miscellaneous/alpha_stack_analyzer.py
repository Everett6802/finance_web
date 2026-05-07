#! /usr/bin/python
# -*- coding: utf8 -*-

import os
import sys
import re
import argparse
import requests  # always safe fallback
try:
	import httpx
	import asyncio
	HAS_HTTPX = True
except ImportError:
	HAS_HTTPX = False
# # import requests
# import httpx
# import asyncio
import pandas as pd
import numpy as np
import time
from datetime import datetime, date, timedelta


class AlphaStackAnalyzer(object):

	DEFAULT_FINMIND_DATE_FORMAT = "%Y-%m-%d"
	DEFAULT_FINMIND_MIN_DATE = date(1900, 1, 1)
	DEFAULT_FINMIND_MIN_DATE_STR = DEFAULT_FINMIND_MIN_DATE.strftime(DEFAULT_FINMIND_DATE_FORMAT)

	FINMIND_BASE_URL = "https://api.finmindtrade.com/api/v4/data"

	@classmethod
	def __build_params(cls, token, dataset_name, stock_symbol, fetch_start=None, fetch_end=None):
		params = {
			"dataset": dataset_name,
			"token": token,
			"data_id": stock_symbol,
		}
		if fetch_start:
			params["start_date"] = fetch_start.strftime("%Y-%m-%d")
		if fetch_end:
			params["end_date"] = fetch_end.strftime("%Y-%m-%d")
		return params


	@classmethod
	def __to_df(cls, data):
		return pd.DataFrame(data["data"])


# ========================
# 基礎抓資料函數
# ========================
	# @classmethod
	# def __get_data(cls, token, dataset_name, stock_symbol, fetch_start=None, fetch_end=None):
	# 	url = "https://api.finmindtrade.com/api/v4/data"
	# 	params = {
	# 		"dataset": dataset_name,
	# 		"token": token,
	# 		"data_id": stock_symbol,
	# 	}
	# 	if fetch_start is not None:
	# 		params["start_date"] = fetch_start.strftime(cls.DEFAULT_FINMIND_DATE_FORMAT)
	# 	else:
	# 		params["start_date"] = cls.DEFAULT_FINMIND_MIN_DATE_STR
	# 	if fetch_end is not None:
	# 		params["end_date"] = fetch_end.strftime(cls.DEFAULT_FINMIND_DATE_FORMAT)

	# 	data = requests.get(url, params=params).json()
	# 	return pd.DataFrame(data["data"])
	@classmethod
	def __get_data(cls, token, dataset_name, stock_symbol, fetch_start=None, fetch_end=None):
		params = cls.__build_params(token, dataset_name, stock_symbol, fetch_start, fetch_end)
		try:
# === 發送 request ===
			if not HAS_HTTPX:
				resp = requests.get(cls.FINMIND_BASE_URL, params=params, timeout=10.0)
			else:
				resp = httpx.get(cls.FINMIND_BASE_URL, params=params, timeout=10.0)
# === HTTP 層錯誤 ===
			resp.raise_for_status()
		except Exception as e:  # === 網路 / HTTP 錯誤 ===
# 這裡同時 cover requests + httpx
			print(f"HTTP ERROR [{stock_symbol}] {dataset_name}: {e}")
			return pd.DataFrame([])
		data = None
		try:
# === JSON parsing ===
			data = resp.json()
		except ValueError:
			print(f"JSON ERROR [{stock_symbol}] {dataset_name}")
			return pd.DataFrame([])
# === API 層 ===
		if data.get("status") != 200:
			print(f"API ERROR [{stock_symbol}] {dataset_name}: {data.get('msg')}")
			return pd.DataFrame([])
		return cls.__to_df(data)


	@classmethod
	async def __get_data_async(cls, client, token, dataset_name, stock_symbol, fetch_start=None, fetch_end=None):
# async 不等於 multi-thread 
# 大量 API request 這是典型：I/O bound
		params = cls.__build_params(token, dataset_name, stock_symbol, fetch_start, fetch_end)
		try:
# === 發送 request ===
			resp = await client.get(cls.FINMIND_BASE_URL, params=params)
# === HTTP 層錯誤 ===
			resp.raise_for_status()
		except Exception as e:  # === 網路 / HTTP 錯誤 ===
			print(f"HTTP ERROR [{stock_symbol}] {dataset_name}: {e}")
			return pd.DataFrame([])
		data = None
		try:
# === JSON parsing ===
			data = resp.json()
		except ValueError:
			print(f"JSON ERROR [{stock_symbol}] {dataset_name}")
			return pd.DataFrame([])
# === API 層 ===
		if data.get("status") != 200:
			print(f"API ERROR [{stock_symbol}] {dataset_name}: {data.get('msg')}")
			return pd.DataFrame([])
		return cls.__to_df(data)


# ========================
# ROE & FCF（財報）
# ========================
	@classmethod
	# def __get_financial_factors(cls, token, stock_symbol, fetch_start=None, fetch_end=None):
	# 	statement = cls.__get_data(token, "TaiwanStockFinancialStatements", stock_symbol, fetch_start, fetch_end)
	def __calculate_financial_factors(cls, statement):
		statement = statement.pivot_table(index="date", columns="type", values="value", aggfunc="first").reset_index()
		statement["ROE"] = statement["NetIncome"] / statement["Equity"]
		statement["FCF"] = statement["CashFlowsFromOperatingActivities"] - statement["CapitalExpenditures"]
		return statement[["date", "ROE", "FCF"]]


# ========================
# PEGY（TTM）
# ========================
	@classmethod
	# def __get_pegy(cls, token, stock_symbol, fetch_start=None, fetch_end=None):
	# 	pe = cls.__get_data(token, "TaiwanStockPER", stock_symbol, fetch_start, fetch_end)
	# 	revenue = cls.__get_data(token, "TaiwanStockMonthRevenue", stock_symbol, fetch_start, fetch_end)
	# 	dividend = cls.__get_data(token, "TaiwanStockDividend", stock_symbol, fetch_start, fetch_end)
	# 	price = cls.__get_data(token, "TaiwanStockPrice", stock_symbol, fetch_start, fetch_end)
	def __calculate_pegy(cls, pe, revenue, dividend, price):
# PE
		# pe = cls.__get_data(token, "TaiwanStockPER", stock_symbol, fetch_start, fetch_end)
		pe = pe.rename(columns={"PER": "PE"})
		pe["date"] = pd.to_datetime(pe["date"])
# 月營收
		# revenue = cls.__get_data(token, "TaiwanStockMonthRevenue", stock_symbol, fetch_start, fetch_end)
		revenue["date"] = pd.to_datetime(revenue["date"])
		revenue = revenue.sort_values("date")
# TTM營收
		revenue["TTM_revenue"] = revenue["revenue"].rolling(12).sum()
		revenue["TTM_YoY"] = revenue["TTM_revenue"].pct_change(12)
# 股利
		# dividend = cls.__get_data(token, "TaiwanStockDividend", stock_symbol, fetch_start, fetch_end)
		dividend["date"] = pd.to_datetime(dividend["date"])
		dividend["cash_dividend"] = dividend["CashDividend"]
# forward fill
		dividend = dividend.sort_values("date")
		dividend["TTM_dividend"] = dividend["cash_dividend"]
		dividend["TTM_dividend"] = dividend["TTM_dividend"].fillna(method="ffill")
# 股價
		# price = cls.__get_data(token, "TaiwanStockPrice", stock_symbol, fetch_start, fetch_end)
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
		df["PEGY"] = df["PE"] / (df["TTM_YoY"] + df["yield"])
		return df[["date", "PEGY"]]


# ========================
# Sharpe Ratio
# ========================
	@classmethod
	# def __get_sharpe(cls, token, stock_symbol, fetch_start=None, fetch_end=None):
	# 	price = cls.__get_data(token, "TaiwanStockPrice", stock_symbol, fetch_start, fetch_end)
	def __calculate_sharpe(cls, price):
		price["date"] = pd.to_datetime(price["date"])
		price = price.sort_values("date")
		price["return"] = price["close"].pct_change()
		window = 252
		price["Sharpe"] = (price["return"].rolling(window).mean() / price["return"].rolling(window).std()) * np.sqrt(252)
		return price[["date", "Sharpe"]]


	def __init__(self, cfg):
		self.xcfg = {
			"finmind_token": None,
			"enable_async": False,
			"stock_symbol_string": None,
			"data_date_range_string": None,
		}
		# import pdb; pdb.set_trace()
		self.xcfg.update(cfg)


	def __enter__(self):
		return self


	def __exit__(self, type, msg, traceback):
		return False


	def __process_data(self, statement, pe, revenue, dividend, price):
# calcuate the paramters
		df1 = self.__calculate_financial_factors(statement)
		df2 = self.__calculate_pegy(pe, revenue, dividend, price)
		df3 = self.__calculate_sharpe(price)
# merge
		df = df2.merge(df1, on="date", how="left")
		df = df.merge(df3, on="date", how="left")
		df = df.sort_values("date")
# 避免未來函數
		df["Sharpe"] = df["Sharpe"].shift(1)
# 清理資料
		df = df.replace([np.inf, -np.inf], np.nan)
		df = df.dropna()
		return df

	
	def __filter_data(self, df):
# 篩選條件（可調整）
		df_filtered = df[(df["PEGY"] < 1) & (df["ROE"] > 0.15) & (df["FCF"] > 0) & (df["Sharpe"] > 0.5)]
		return df_filtered


	# ========================
	# 單支分析（同步）
	# ========================
	def __analyze(self, stock_symbol, fetch_start=None, fetch_end=None):
# get data 
		statement = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockFinancialStatements", stock_symbol, fetch_start, fetch_end)
		pe = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockPER", stock_symbol, fetch_start, fetch_end)
		revenue = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockMonthRevenue", stock_symbol, fetch_start, fetch_end)
		dividend = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockDividend", stock_symbol, fetch_start, fetch_end)
		price = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockPrice", stock_symbol, fetch_start, fetch_end)
# # calcuate the paramters
# 		df1 = self.__calculate_financial_factors(statement)
# 		df2 = self.__calculate_pegy(pe, revenue, dividend, price)
# 		df3 = self.__calculate_sharpe(price)
# # merge
# 		df = df2.merge(df1, on="date", how="left")
# 		df = df.merge(df3, on="date", how="left")
# 		df = df.sort_values("date")
# # 避免未來函數
# 		df["Sharpe"] = df["Sharpe"].shift(1)
# # 清理資料
# 		df = df.replace([np.inf, -np.inf], np.nan)
# 		df = df.dropna()
		df = self.__process_data(statement, pe, revenue, dividend, price)
		df_filtered = self.__filter_data(df)
		# print(df_filtered.tail())
		# df.to_csv("factor_data.csv", index=False)
		return df_filtered


	# ========================
	# 多支分析（async）
	# ========================
# async 是Python 保留字
	async def __analyze_many(self, stock_symbol_list, fetch_start=None, fetch_end=None):
# 同時分析很多股票->每支股票同時打API->等待時切去別支->最後收集結果
		if not HAS_HTTPX:
			raise ImportError("The httpx package is NOT installed")
		results = []
		limits = httpx.Limits(max_connections=10)
# 建立一個「非同步 HTTP session
		async with httpx.AsyncClient(limits=limits, timeout=10.0) as client:
# 限制同時數量（Semaphore）:同一時間最多 5 個 task 在跑 
			sem = asyncio.Semaphore(5)
			async def task(stock_symbol):
# 進入 semaphore 拿到「執行許可」 如果已經 5 個人在跑：這個 task 會等待。
				async with sem:
					try:
# get data 
# await 是什麼？「這段 I/O 等待時，先去做別的 task」
# 同步：等 API 回應 → CPU 發呆 非同步：等 API 回應時 → 去抓別支股票
						statement = await self.__get_data_async(client, self.xcfg["finmind_token"], "TaiwanStockFinancialStatements", stock_symbol, fetch_start, fetch_end)
						pe = await self.__get_data_async(client, self.xcfg["finmind_token"], "TaiwanStockPER", stock_symbol, fetch_start, fetch_end)
						revenue = await self.__get_data_async(client, self.xcfg["finmind_token"], "TaiwanStockMonthRevenue", stock_symbol, fetch_start, fetch_end)
						dividend = await self.__get_data_async(client, self.xcfg["finmind_token"], "TaiwanStockDividend", stock_symbol, fetch_start, fetch_end)
						price = await self.__get_data_async(client, self.xcfg["finmind_token"], "TaiwanStockPrice", stock_symbol, fetch_start, fetch_end)
						df = self.__process_data(statement, pe, revenue, dividend, price)
						df_filtered = self.__filter_data(df)
						return df_filtered
					except Exception as e:
						print(f"error {stock_symbol}: {e}")
						return None
# 建立所有 tasks
			tasks = [task(stock_symbol) for stock_symbol in stock_symbol_list]
# 這會得到：[df_2330, df_2317, df_2454, ...] 每個都是：pandas.DataFrame
			results = await asyncio.gather(*tasks)  # 真正開始跑
# pd.concat把多個 DataFrame「接起來」
		'''
		df1
		stock	PEGY
		2330	0.8
		df2
		stock	PEGY
		2317	0.9
		pd.concat([df1, df2]) 變：
		stock	PEGY
		2330	0.8
		2317	0.9
		'''
		df_list = [r for r in results if r is not None]
		if len(df_list) == 0:
			return pd.DataFrame()
		return pd.concat(df_list, ignore_index=True)


	def analyze(self):
		if self.xcfg["stock_symbol_string"] is None:
			print("Warning: No stock to fetch...")
			return
		stock_symbol_list = self.xcfg["stock_symbol_string"].split(",")
		date_range_start_str = date_range_end_str = None
		if self.xcfg["data_date_range_string"] is not None:
			date_range_elems = self.xcfg["data_date_range_string"].split(":")
			if len(date_range_elems) == 2:
				if date_range_elems[0] != "":
					date_range_start_str = date_range_elems[0]
				if date_range_elems[1] != "":
					date_range_end_str = date_range_elems[1]
			else:
				print("Error: Incorrect date range format[%s]" % self.xcfg["data_date_range_string"])
				return 
		if self.xcfg["enable_async"] and not HAS_HTTPX:
			print("Warning: No httpx pakcage, fallback to the synchronous mode")
			self.xcfg["enable_async"] = False
		if self.xcfg["enable_async"]:
			asyncio.run(self.__analyze_many(stock_symbol_list, date_range_start_str, date_range_end_str))
		else:
			for stock_symbol in stock_symbol_list:
				self.__analyze(stock_symbol, date_range_start_str, date_range_end_str)


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
	parser.add_argument('-a', '--analyze', required=False, action='store_true', help='Analysz data and exit.')
	parser.add_argument('--finmind_token', required=False, help='The FinMind Token')
	parser.add_argument('--enable_async', required=False, help='Fetch data in the asynchronous mode')
	parser.add_argument('--stock_symbol_list', required=False, help='The stock symbol list. Mutiple stock symbols are split by comma. Ex: 00850.TW,00881.TW,00692.TW,MSFT,GOOG')
	args = parser.parse_args()
	# import pdb; pdb.set_trace()
	cfg = {}
	if args.finmind_token is not None: cfg['finmind_token'] = args.finmind_token
	if args.enable_async is not None: cfg['enable_async'] = True
	# import pdb; pdb.set_trace()
	with AlphaStackAnalyzer(cfg) as obj:
		if args.analyze:
			obj.analyze()
			sys.exit(0)