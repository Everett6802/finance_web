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
		if fetch_start is not None:
			if isinstance(fetch_start, str):
				if re.match(r"\d{4}-\d{2}-\d{2}", fetch_start) is None:
					print(f"Error: Incorrect start date format[{fetch_start}]")
					raise ValueError(f"Incorrect start date format[{fetch_start}]")
				params["start_date"] = fetch_start
			else:
				params["start_date"] = fetch_start.strftime("%Y-%m-%d")
		if fetch_end:
			if isinstance(fetch_end, str):
				if re.match(r"\d{4}-\d{2}-\d{2}", fetch_end) is None:
					print(f"Error: Incorrect end date format[{fetch_end}]")
					raise ValueError(f"Incorrect end date format[{fetch_end}]")
				params["end_date"] = fetch_end
			else:
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
	def __calculate_financial_factors(cls, income_stmt, balance_sheet, cashflow):

		def pivot(df):
			if df.empty:
				return pd.DataFrame()

			df = df.pivot_table(
				index="date",
				columns="type",
				values="value",
				aggfunc="first"
			).reset_index()

			df["date"] = pd.to_datetime(df["date"])
			return df

		# ==================================================
		# 1. Pivot 原始財報
		# ==================================================
		income_stmt = pivot(income_stmt)
		balance_sheet = pivot(balance_sheet)
		cashflow = pivot(cashflow)

		# ==================================================
		# 2. 找出財報欄位
		# ==================================================

		# ---- Net Income ----
		net_income_col = next(
			(c for c in [
				"NetIncome",
				"IncomeAfterTaxes",
				"ProfitLoss"
			] if c in income_stmt.columns),
			None
		)

		if net_income_col is None:
			raise ValueError("No Net Income column")

		# ---- Equity ----
		equity_col = next(
			(c for c in [
				"EquityAttributableToOwnersOfParent",
				"Equity",
				"TotalEquity"
			] if c in balance_sheet.columns),
			None
		)

		if equity_col is None:
			raise ValueError("No Equity column")

		# ---- Operating Cash Flow ----
		ocf_col = next(
			(c for c in [
				"CashFlowsFromOperatingActivities",
				"NetCashProvidedByUsedInOperatingActivities"
			] if c in cashflow.columns),
			None
		)

		if ocf_col is None:
			raise ValueError("No OCF column")

		# ---- Capital Expenditure ----
		capex_col = next(
			(c for c in [
				"CapitalExpenditures",
				"PurchaseOfPropertyPlantAndEquipment",
				"PropertyAndPlantAndEquipment"
			] if c in cashflow.columns),
			None
		)

		if capex_col is None:
			# cashflow["CAPEX"] = 0
			# capex_col = "CAPEX"
			raise ValueError(
				"No CAPEX column found in Cash Flow Statement"
			)

		# ==================================================
		# 3. Income Statement：單季淨利 → 全年淨利
		# ==================================================
		income_stmt = income_stmt.sort_values("date").copy()

		income_stmt["year"] = income_stmt["date"].dt.year
		income_stmt["quarter"] = income_stmt["date"].dt.quarter

		# 每年必須具備完整 Q1 ~ Q4
		quarter_count = (
			income_stmt.groupby("year")["quarter"]
			.transform("nunique")
		)

		income_stmt["AnnualNetIncome"] = (
			income_stmt.groupby("year")[net_income_col]
			.transform(
				lambda s: s.sum(min_count=4)
			)
		)

		# 缺少任一季度時，不計算全年淨利
		income_stmt.loc[
			quarter_count != 4,
			"AnnualNetIncome"
		] = np.nan

		# 只保留年底資料
		income_stmt = income_stmt[
			income_stmt["date"].dt.month == 12
		].copy()

		# ==================================================
		# 4. Balance Sheet：保留年底股東權益
		# ==================================================
		balance_sheet = balance_sheet[
			balance_sheet["date"].dt.month == 12
		].copy()

		balance_sheet = balance_sheet.sort_values("date")

		# 計算期初、期末平均股東權益
		balance_sheet["BeginningEquity"] = (
			balance_sheet[equity_col].shift(1)
		)

		balance_sheet["AverageEquity"] = (
			balance_sheet["BeginningEquity"]
			+ balance_sheet[equity_col]
		) / 2

		# ==================================================
		# 5. Cash Flow：暫時維持原本邏輯
		# ==================================================
		# print("\n[DEBUG] 現金流量表欄位")
		# print("OCF column  :", ocf_col)
		# print("CAPEX column:", capex_col)

		# print("\n[DEBUG] 2025 年各季現金流量原始資料")

		# debug_cashflow = cashflow[
		# 	cashflow["date"].dt.year == 2025
		# ].copy()

		# print(
		# 	debug_cashflow[
		# 		["date", ocf_col, capex_col]
		# 	].to_string(index=False)
		# )
		cashflow = cashflow[
			cashflow["date"].dt.month == 12
		].copy()

		# ==================================================
		# 6. Merge 年度財報
		# ==================================================
		df = income_stmt[
			["date", "AnnualNetIncome"]
		].merge(
			balance_sheet[
				[
					"date",
					equity_col,
					"BeginningEquity",
					"AverageEquity"
				]
			],
			on="date",
			how="inner"
		).merge(
			cashflow[
				["date", ocf_col, capex_col]
			],
			on="date",
			how="left"
		)

		df = df.sort_values("date")

		# ==================================================
		# 7. ROE
		# ==================================================
		df["ROE"] = np.where(
			df["AverageEquity"] > 0,
			df["AnnualNetIncome"] / df["AverageEquity"],
			np.nan
		)

		# ==================================================
		# 8. FCF（暫時沿用原本計算方式）
		# ==================================================
		df["FCF"] = (
			df[ocf_col] - df[capex_col].abs()
		)

		# ==================================================
		# 9. 清理異常值
		# ==================================================
		df["ROE"] = df["ROE"].replace(
			[np.inf, -np.inf],
			np.nan
		)

		df["FCF"] = df["FCF"].replace(
			[np.inf, -np.inf],
			np.nan
		)

		# # ==================================================
		# # 10. DEBUG
		# # ==================================================
		# print("\n[DEBUG] 年度 ROE 計算資料")

		# debug_df = df[
		# 	[
		# 		"date",
		# 		"AnnualNetIncome",
		# 		"BeginningEquity",
		# 		equity_col,
		# 		"AverageEquity",
		# 		"ROE"
		# 	]
		# ].tail(10).copy()

		# print(debug_df.to_string(index=False))

		# ==================================================
		# 11. Return
		# ==================================================
		return df[["date", "ROE", "FCF"]]


# ========================
# Dividend（TTM）
# ========================
	@classmethod
	def __calculate_ttm_dividend(cls, dividend_df, lookback_days=900, outlier_multiple=3):
		"""
		自動判斷配息頻率並計算 TTM Dividend
		1. 先觀察過去數年的配息模式
		2. 自動判斷： 月配 季配 半年配 年配
		3. 檢查配息週期是否穩定 若穩定：使用最近 N 次配息加總 若不穩定： fallback → rolling(365D)
		4. 過濾特別股利（outlier）
		* Parameters
		dividend_df : pd.DataFrame   必須包含: date cash_dividend
		lookback_days : int   用於分析配息模式的歷史天數
		outlier_multiple : float   超過 median * N 視為特別股利
		* Returns
		result : dict   {"TTM_dividend": float, "frequency": str, "stable": bool, "warning": str or None}
		"""
		df = dividend_df.copy()
		if len(df) == 0:
			return {
				"TTM_dividend": np.nan,
				"frequency": "unknown",
				"stable": False,
				"warning": "No dividend data",
			}
		df["date"] = pd.to_datetime(df["date"])
		df = df.sort_values("date")
# 只看近幾年
		latest_date = df["date"].max()
		cutoff_date = latest_date - pd.Timedelta(days=lookback_days)
		df = df[df["date"] >= cutoff_date].copy()
# 至少需要兩筆資料
		if len(df) < 2:
			ttm = df["cash_dividend"].sum()
			return {
				"TTM_dividend": ttm,
				"frequency": "unknown",
				"stable": False,
				"warning": "Not enough dividend history",
			}
# 計算間隔
# df["date"].diff() => 計算「與前一筆的差值」但第一筆「沒有前一筆可以比較」=> 要dropna()才能算平均與標準差
		df["interval_days"] = df["date"].diff().dt.days
		intervals = df["interval_days"].dropna()
		mean_days = intervals.mean()
		std_days = intervals.std()
# 判斷配息頻率
		if mean_days <= 45:
			frequency = "monthly"
			expected_count = 12
			tolerance_std = 5
		elif mean_days <= 120:
			frequency = "quarterly"
			expected_count = 4
			tolerance_std = 15
		elif mean_days <= 240:
			frequency = "semiannual"
			expected_count = 2
			tolerance_std = 30
		else:
			frequency = "annual"
			expected_count = 1
			tolerance_std = 45
# 判斷穩定性
		stable = std_days <= tolerance_std
# 過濾特別股利
		median_dividend = df["cash_dividend"].median()
		upper_limit = median_dividend * outlier_multiple
# series.clip(lower=?, upper=?)
# 作用：lower	最小值限制 upper 最大值限制
# s = pd.Series([1, 2, 100])
# s.clip(upper=10) -> [1, 2, 10]
		df["filtered_dividend"] = df["cash_dividend"].clip(upper=upper_limit)
		if stable and len(df) >= expected_count:  # 穩定 → 最近 N 次加總
			ttm_dividend = df["filtered_dividend"].tail(expected_count).sum()
			warning = None
		else:  # 不穩定 → rolling(365D)
			warning = (f"Dividend schedule unstable (std={std_days:.1f} days). Fallback to rolling 365D calculation.")
			ttm_dividend = df.set_index("date")["filtered_dividend"].rolling("365D").sum().iloc[-1]
		return {
			"TTM_dividend": float(ttm_dividend),
			"frequency": frequency,
			"stable": stable,
			"warning": warning,
		}


	@classmethod
	def __convert_cumulative_to_quarterly(cls, df, eps_col):

		df = df.copy()

		df["date"] = pd.to_datetime(df["date"])

		df = df.sort_values("date")

		quarterly_eps = []

		prev_year = None
		prev_cum_eps = 0.0

		for _, row in df.iterrows():

			current_date = row["date"]
			current_year = current_date.year

			cum_eps = row[eps_col]

			if pd.isna(cum_eps):
				quarterly_eps.append(np.nan)
				continue

			# 新年度 Q1
			if prev_year != current_year:
				q_eps = cum_eps
			else:
				q_eps = cum_eps - prev_cum_eps

			quarterly_eps.append(q_eps)

			prev_cum_eps = cum_eps
			prev_year = current_year

		df["quarter_eps"] = quarterly_eps

		return df

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

		# ---- PE ----
		pe = pe.rename(columns={"PER": "PE"}).copy()
		pe["date"] = pd.to_datetime(pe["date"])
		pe = pe.sort_values("date")

		# ---- Revenue: TTM revenue growth ----
		revenue = revenue.copy()
		revenue["date"] = pd.to_datetime(revenue["date"])
		revenue = revenue.sort_values("date")

		# 每月營收 → TTM 營收
		revenue["TTM_revenue"] = revenue["revenue"].rolling(
			window=12,
			min_periods=12
		).sum()

		# TTM YoY
		revenue["TTM_YoY"] = revenue["TTM_revenue"].pct_change(
			periods=12,
			fill_method=None
		)

		# ==================================================
		# Revenue available date
		# ==================================================

		REVENUE_FALLBACK_DAY = 16

		# 部分歷史資料可能沒有 create_time 欄位
		if "create_time" not in revenue.columns:
			revenue["create_time"] = pd.NaT

		revenue["create_time"] = pd.to_datetime(
			revenue["create_time"],
			errors="coerce"
		)

		# FinMind 的 date 已經是營收所屬月份的次月 1 日
		#
		# 例如：
		# revenue_year  = 2025
		# revenue_month = 12
		# date          = 2026-01-01
		# fallback_date = 2026-01-16

		revenue["fallback_date"] = (
			revenue["date"].dt.to_period("M").dt.to_timestamp()
			+ pd.Timedelta(days=REVENUE_FALLBACK_DAY - 1)
		)

		# 有 create_time 時，從隔天開始使用
		revenue["available_date"] = (
			revenue["create_time"].dt.normalize()
			+ pd.Timedelta(days=1)
		)

		# 記錄原始日期來源
		revenue["available_date_source"] = np.where(
			revenue["create_time"].notna(),
			"create_time",
			"fallback"
		)

		# 缺少 create_time 時，使用 fallback_date
		revenue["available_date"] = (
			revenue["available_date"]
			.fillna(revenue["fallback_date"])
		)

		# 避免異常的 create_time 早於營收所屬月份結束
		invalid_available_date = (
			revenue["available_date"] < revenue["date"]
		)

		revenue.loc[
			invalid_available_date,
			"available_date"
		] = revenue.loc[
			invalid_available_date,
			"fallback_date"
		]

		revenue.loc[
			invalid_available_date,
			"available_date_source"
		] = "fallback_invalid_create_time"

		# 保留 cummax 調整前的日期，供 DEBUG 使用
		revenue["original_available_date"] = (
			revenue["available_date"]
		)

		# 確保較新的 TTM 營收不會早於
		# 較舊營收資料的可用日期
		revenue["available_date"] = (
			revenue["available_date"].cummax()
		)

		# 記錄因 cummax 而被延後的資料
		revenue["available_date_adjusted"] = (
			revenue["available_date"]
			!= revenue["original_available_date"]
		)

		# ==================================================
		# DEBUG A: Historical revenue availability
		# ==================================================

		print("\n[DEBUG A] 月營收歷史資料")

		print(
			revenue[
				[
					"date",
					"revenue_year",
					"revenue_month",
					"create_time",
					"fallback_date",
					"available_date",
					"available_date_source",
					"available_date_adjusted",
					"TTM_revenue",
					"TTM_YoY"
				]
			].to_string(index=False)
		)

		print("\n[DEBUG A] Revenue availability statistics")

		print("Total rows:", len(revenue))

		print(
			"Missing create_time:",
			revenue["create_time"].isna().sum()
		)

		print(
			"Missing available_date:",
			revenue["available_date"].isna().sum()
		)

		print(
			"Adjusted available_date:",
			revenue["available_date_adjusted"].sum()
		)

		print("\n[DEBUG A] Available date source count")

		print(
			revenue["available_date_source"]
			.value_counts()
			.to_string()
		)

		# ==================================================
		# DEBUG B: create_time delay statistics
		# ==================================================

		revenue["create_time_delay_days"] = (
			revenue["create_time"].dt.normalize()
			- revenue["date"].dt.normalize()
		).dt.days

		print("\n[DEBUG B] create_time delay statistics")

		print(
			revenue.groupby("revenue_year")[
				"create_time_delay_days"
			].agg(
				["size", "count", "min", "median", "max"]
			).to_string()
		)

		# ---- Prepare revenue growth ----

		revenue_growth = revenue[
			[
				"available_date",
				"available_date_source",
				"revenue_year",
				"revenue_month",
				"TTM_revenue",
				"TTM_YoY"
			]
		].dropna().sort_values(
			[
				"available_date",
				"revenue_year",
				"revenue_month"
			]
		)

		# 記錄移除重複日期之前的筆數
		revenue_growth_rows_before = len(revenue_growth)

		# 同一天有多筆營收資料可用時，
		# 保留營收所屬月份最新的一筆
		revenue_growth = revenue_growth.drop_duplicates(
			subset=["available_date"],
			keep="last"
		).reset_index(drop=True)

		# ==================================================
		# DEBUG C: Revenue growth data coverage
		# ==================================================

		print("\n[DEBUG C] revenue_growth date range")

		print("Rows :", len(revenue_growth))
		print("Start:", revenue_growth["available_date"].min())
		print("End  :", revenue_growth["available_date"].max())

		print(
			"Removed duplicate available_date:",
			revenue_growth_rows_before - len(revenue_growth)
		)

		print("\n[DEBUG C] revenue_growth yearly count")

		print(
			revenue_growth.groupby(
				revenue_growth["available_date"].dt.year
			).size().to_string()
		)

		# ---- Dividend ----

		dividend = dividend.copy()
		dividend["date"] = pd.to_datetime(dividend["date"])

		dividend_cols = [
			"CashEarningsDistribution",
			"CashStatutorySurplus"
		]

		if not any(col in dividend.columns for col in dividend_cols):
			raise ValueError("No cash dividend columns found")

		for col in dividend_cols:

			if col not in dividend.columns:
				dividend[col] = 0.0

			dividend[col] = pd.to_numeric(
				dividend[col],
				errors="coerce"
			).fillna(0.0)

		dividend["cash_dividend"] = (
			dividend["CashEarningsDistribution"]
			+ dividend["CashStatutorySurplus"]
		)

		dividend = dividend.sort_values("date")

		# ---- Price ----

		price = price.copy()
		price["date"] = pd.to_datetime(price["date"])
		price = price.sort_values("date")

		# ---- Merge: PE and Revenue ----

		df = pe[["date", "PE"]].copy()

		df = pd.merge_asof(
			df.sort_values("date"),
			revenue_growth,
			left_on="date",
			right_on="available_date",
			direction="backward",
			allow_exact_matches=True
		)

		# ==================================================
		# TTM Dividend: trailing 365 calendar days
		# ==================================================

		# 將同一天的現金股利加總
		dividend_daily = (
			dividend.groupby("date")["cash_dividend"]
			.sum(min_count=1)
			.sort_index()
		)

		# 建立包含交易日及股利事件日的完整時間軸
		timeline = dividend_daily.index.union(
			pd.DatetimeIndex(df["date"])
		).unique().sort_values()

		# 非股利事件日的現金股利設為 0
		dividend_daily = dividend_daily.reindex(
			timeline,
			fill_value=0.0
		)

		# 每一天往前 365 天的現金股利總和
		ttm_dividend = dividend_daily.rolling(
			window="365D",
			min_periods=1,
			closed="right"
		).sum()

		# 將 TTM 股利對應回每個交易日
		df["TTM_dividend"] = df["date"].map(
			ttm_dividend
		)

		# 保留股利事件日的 TTM 股利，供 DEBUG 使用
		dividend["TTM_dividend"] = dividend["date"].map(
			ttm_dividend
		)

		# ---- Merge: Price ----

		df = pd.merge_asof(
			df,
			price[["date", "close"]],
			on="date",
			direction="backward"
		)

		# ---- Dividend yield ----

		df["yield"] = df["TTM_dividend"] / df["close"]

		# ==================================================
		# PEGY
		# ==================================================

		# TTM_YoY 是小數，例如：
		# 0.403 = 40.3%
		#
		# PEGY 使用百分點：
		# 40.3 + 0.96 = 41.26

		growth_pct = df["TTM_YoY"] * 100
		yield_pct = df["yield"] * 100

		denom = growth_pct + yield_pct

		df["PEGY"] = np.where(
			(denom > 0)
			& (df["PE"] > 0)
			& (df["close"] > 0),
			df["PE"] / denom,
			np.nan
		)

		# ---- Remove invalid values ----

		df["PEGY"] = df["PEGY"].replace(
			[np.inf, -np.inf],
			np.nan
		)

		# ==================================================
		# DEBUG D: PEGY yearly data coverage
		# ==================================================

		print("\n[DEBUG D] PEGY yearly data coverage")

		print(
			df.groupby(df["date"].dt.year).agg(
				total_days=("date", "size"),
				valid_pe_days=("PE", "count"),
				valid_revenue_days=("TTM_YoY", "count"),
				valid_dividend_days=("TTM_dividend", "count"),
				valid_price_days=("close", "count"),
				valid_pegy_days=("PEGY", "count")
			).to_string()
		)

		print("\n[DEBUG D] PEGY invalid reason count")

		print("Total rows:", len(df))

		print(
			"Missing PE:",
			df["PE"].isna().sum()
		)

		print(
			"Missing TTM_YoY:",
			df["TTM_YoY"].isna().sum()
		)

		print(
			"Missing TTM_dividend:",
			df["TTM_dividend"].isna().sum()
		)

		print(
			"Missing close:",
			df["close"].isna().sum()
		)

		print(
			"PE <= 0:",
			(df["PE"] <= 0).sum()
		)

		print(
			"Denominator <= 0:",
			(denom <= 0).sum()
		)

		print(
			"Close <= 0:",
			(df["close"] <= 0).sum()
		)

		print(
			"Invalid PEGY:",
			df["PEGY"].isna().sum()
		)

		valid_pegy = df[df["PEGY"].notna()]

		print("\n[DEBUG D] Valid PEGY date range")

		print("Rows :", len(valid_pegy))
		print("Start:", valid_pegy["date"].min())
		print("End  :", valid_pegy["date"].max())

		# ==================================================
		# DEBUG E: Check look-ahead bias
		# ==================================================

		future_revenue_count = (
			df["available_date"] > df["date"]
		).sum()

		print("\n[DEBUG E] Revenue date alignment check")

		print(
			"Future revenue records:",
			future_revenue_count
		)

		# 理論上 merge_asof backward 不應發生此情況
		if future_revenue_count > 0:
			raise ValueError(
				"Revenue available_date is later than trading date"
			)

		# ==================================================
		# DEBUG F: Latest PEGY data
		# ==================================================

		print("\n[DEBUG F] PEGY 最新計算資料")

		debug_cols = [
			"date",
			"available_date",
			"available_date_source",
			"revenue_year",
			"revenue_month",
			"PE",
			"TTM_revenue",
			"TTM_YoY",
			"TTM_dividend",
			"close",
			"yield",
			"PEGY"
		]

		print(
			df[debug_cols]
			.tail(25)
			.to_string(index=False)
		)

		# ---- DEBUG: Revenue ----

		print("\n[DEBUG] 月營收欄位")

		print(revenue.columns.tolist())

		print("\n[DEBUG] 月營收完整資料")

		print(
			revenue.tail(6).to_string(index=False)
		)

		# ---- DEBUG: Dividend ----

		print("\n[DEBUG] 股利原始日期")

		print(
			dividend[
				[
					"date",
					"cash_dividend",
					"TTM_dividend"
				]
			].tail(8).to_string(index=False)
		)

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

		mean = price["return"].rolling(window).mean()
		std = price["return"].rolling(window).std()

		price["Sharpe"] = (mean / std) * np.sqrt(252)

		price["Sharpe"] = price["Sharpe"].replace([np.inf, -np.inf], np.nan)

		price["Sharpe"] = price["Sharpe"].fillna(0)
		price["Sharpe"] = price["Sharpe"].clip(-5, 5)

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


	def __process_data(
		self, income_stmt, balance_sheet, cashflow,
		pe, revenue, dividend, price
	):
		# ==================================================
		# 1. 檢查原始資料
		# ==================================================
		datasets = {
			"Income Statement": income_stmt,
			"Balance Sheet": balance_sheet,
			"Cash Flow": cashflow,
			"PE": pe,
			"Revenue": revenue,
			"Dividend": dividend,
			"Price": price,
		}

		empty_names = [
			name for name, data in datasets.items()
			if data is None or data.empty
		]

		if empty_names:
			print(
				"Warning: The following dataframes are empty: "
				+ ", ".join(empty_names)
			)
			return None

		# ==================================================
		# 2. 計算三組指標
		# ==================================================
		df1 = self.__calculate_financial_factors(
			income_stmt, balance_sheet, cashflow
		)
		df2 = self.__calculate_pegy(
			pe, revenue, dividend, price
		)
		df3 = self.__calculate_sharpe(price)

		if any(x is None or x.empty for x in [df1, df2, df3]):
			print("Warning: Factor calculation returned empty data")
			return None

		# ==================================================
		# 3. 年度財報的保守可用日期
		# ==================================================
		df1 = df1.copy()
		df1["date"] = pd.to_datetime(df1["date"])

		# 財報年度 Y，於 Y+1 年 4 月 1 日起使用
		# 注意：這是保守代理日期，不是實際公告日期
		df1["available_date"] = pd.to_datetime(
			(df1["date"].dt.year + 1).astype(str)
			+ "-04-01"
		)

		# 2020 年因缺少 2019 年底權益，ROE 為 NaN
		# 排除無法計算的年度財報
		df1 = df1.dropna(
			subset=["ROE", "FCF", "available_date"]
		)

		df1 = (
			df1.sort_values("available_date")
			.drop_duplicates(
				subset=["available_date"],
				keep="last"
			)
		)

		# ==================================================
		# 4. 準備每日指標
		# ==================================================
		df2 = df2.copy()
		df3 = df3.copy()

		df2["date"] = pd.to_datetime(df2["date"])
		df3["date"] = pd.to_datetime(df3["date"])

		df2 = df2.sort_values("date")
		df3 = df3.sort_values("date")

		# ==================================================
		# 5. 合併年度財報與每日 PEGY
		# ==================================================
		df = pd.merge_asof(
			df2,
			df1[
				["available_date", "ROE", "FCF"]
			],
			left_on="date",
			right_on="available_date",
			direction="backward",
			allow_exact_matches=True
		)

		# ==================================================
		# 6. 合併 Sharpe
		# ==================================================
		df = df.merge(
			df3[["date", "Sharpe"]],
			on="date",
			how="left",
			validate="one_to_one"
		)

		df = df.sort_values("date").reset_index(drop=True)

		# 當日收盤後計算的 Sharpe，
		# 最早只能用於下一交易日的訊號
		df["Sharpe"] = df["Sharpe"].shift(1)

		# ==================================================
		# 7. 清理資料
		# ==================================================
		df = df.replace(
			[np.inf, -np.inf],
			np.nan
		)

		df = df.dropna(
			subset=["PEGY", "ROE", "FCF", "Sharpe"]
		).copy()

		# # ==================================================
		# # 8. DEBUG
		# # ==================================================
		# print("\n[DEBUG] 年度財報可用日期")
		# print(
		# 	df1[
		# 		["date", "available_date", "ROE", "FCF"]
		# 	].to_string(index=False)
		# )

		# print("\n[DEBUG] 最終資料範圍")

		# if not df.empty:
		# 	print("Start:", df["date"].min())
		# 	print("End  :", df["date"].max())
		# 	print("Rows :", len(df))

		# 	print("\n[DEBUG] 年度資料筆數")
		# 	print(
		# 		df.groupby(df["date"].dt.year)
		# 		.size()
		# 		.to_string()
		# 	)
		# else:
		# 	print("No valid records")

		# ==================================================
		# 9. Return
		# ==================================================
		return df[["date", "PEGY", "ROE", "FCF", "Sharpe"]]

	
	def __filter_data(self, df):
# 篩選條件（可調整）
		# df_filtered = df[(df["PEGY"] < 1) & (df["ROE"] > 0.15) & (df["FCF"] > 0) & (df["Sharpe"] > 0.5)]
		df_filtered = df[(df["PEGY"] < 20) & (df["ROE"] > 0.08) & (df["FCF"] > 0) & (df["Sharpe"] > 0)]
		return df_filtered


	# ========================
	# 單支分析（同步）
	# ========================
	def __analyze(self, stock_symbol, fetch_start=None, fetch_end=None):
# get data 
		income_stmt = self.__get_data(
			self.xcfg["finmind_token"],
			"TaiwanStockFinancialStatements",
			stock_symbol,
			fetch_start,
			fetch_end,
		)
		# print(f"[DEBUG] Income Statement data: {income_stmt}")

		balance_sheet = self.__get_data(
			self.xcfg["finmind_token"],
			"TaiwanStockBalanceSheet",
			stock_symbol,
			fetch_start,
			fetch_end,
		)

		cashflow = self.__get_data(
			self.xcfg["finmind_token"],
			"TaiwanStockCashFlowsStatement",
			stock_symbol,
			fetch_start,
			fetch_end,
		)
		pe = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockPER", stock_symbol, fetch_start, fetch_end)
		revenue = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockMonthRevenue", stock_symbol, fetch_start, fetch_end)
		income_stmt = self.__get_data(self.xcfg["finmind_token"], "TaiwanStockFinancialStatements", stock_symbol, fetch_start, fetch_end)
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
		# import pdb; pdb.set_trace()
		df = self.__process_data(income_stmt, balance_sheet, cashflow, pe, revenue, dividend, price)
		if df is None:
			print(f"Error: Failed to process data for {stock_symbol}")
			return
		df_filtered = self.__filter_data(df)
		df_filtered["stock_symbol"] = stock_symbol
		# print(df_filtered.tail())
		# df.to_csv("factor_data.csv", index=False)
		self.print_analysis(df, df_filtered, stock_symbol)
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
						if df is None:
							print(f"Error: Failed to process data for {stock_symbol}")
							return None
						df_filtered = self.__filter_data(df)
						df_filtered["stock_symbol"] = stock_symbol
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
		df_targets = None
		if self.xcfg["enable_async"]:
			df_targets = asyncio.run(self.__analyze_many(stock_symbol_list, date_range_start_str, date_range_end_str))
		else:
			# import pdb; pdb.set_trace()
			df_list = []
			for stock_symbol in stock_symbol_list:
				result = self.__analyze(stock_symbol, date_range_start_str, date_range_end_str)
				if result is not None:
					df_list.append(result)
			if len(df_list) == 0:
				return pd.DataFrame()
			df_targets = pd.concat(df_list, ignore_index=True)
		return df_targets


	def print_analysis(self, df_raw, df_filtered, stock_symbol):
		"""
		印出分析結果：原始計算數據 + 篩選後結果
		df_raw     : __process_data 回傳的完整 DataFrame
		df_filtered: __filter_data  回傳的篩選後 DataFrame
		stock_symbol: 股票代號（用於標題）
		"""
		SEP = "=" * 70

		print(f"\n{SEP}")
		print(f"  股票代號: {stock_symbol}")
		print(SEP)

		# ── 最新一筆各指標數值 ──────────────────────────────────────
		if df_raw is not None and not df_raw.empty:
			latest = df_raw.iloc[-1]
			print(f"  【最新數據】 ({latest['date'].date()})")
			print(f"    PEGY   : {latest['PEGY']:.4f}")
			print(f"    ROE    : {latest['ROE']:.2%}")
			print(f"    FCF    : {latest['FCF']:>15,.0f}")
			print(f"    Sharpe : {latest['Sharpe']:.4f}")
		else:
			print("  【最新數據】 無資料")

		print()

		# ── 篩選條件說明 ────────────────────────────────────────────
		print("  【篩選條件】")
		# print("    PEGY < 1  |  ROE > 15%  |  FCF > 0  |  Sharpe > 0.5")
		print("    PEGY < 20  |  ROE > 8%  |  FCF > 0  |  Sharpe > 0")
		print()

		# ── 篩選後結果 ──────────────────────────────────────────────
		if df_filtered is not None and not df_filtered.empty:
			total_days  = len(df_raw) if df_raw is not None else 0
			passed_days = len(df_filtered)
			hit_rate    = passed_days / total_days if total_days > 0 else 0.0

			print(f"  【符合篩選條件的交易日】 {passed_days} / {total_days} 天 ({hit_rate:.1%})")
			print()

			# 取最近 10 筆
			display = df_filtered.tail(10).copy()
			display["date"] = display["date"].dt.date
			display["ROE"]  = display["ROE"].map(lambda x: f"{x:.2%}")
			display["FCF"]  = display["FCF"].map(lambda x: f"{x:,.0f}")
			display["PEGY"] = display["PEGY"].map(lambda x: f"{x:.4f}")
			display["Sharpe"] = display["Sharpe"].map(lambda x: f"{x:.4f}")

			cols = ["date", "PEGY", "ROE", "FCF", "Sharpe"]
			print(display[cols].to_string(index=False))
		else:
			print("  【符合篩選條件的交易日】 0 天（無符合紀錄）")

		print(f"{SEP}\n", flush=True)

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
	parser.add_argument('--stock_symbol_list', required=False, help='The Taiwan stock symbol list. Mutiple stock symbols are split by comma. Ex: 00850,00881,00692')
	parser.add_argument('--data_date_range', required=False, 
		 help='''The data during the date range.
  Date range:
    Format: yy1-mm1-dd1:yy2-mm2-dd2   From yy1-mm1-dd1 to yy2-mm2-dd2   Ex: 2014-09-04:2025-10-15
    Format: yy-mm-dd:   From yy-mm-dd to 'the last date of the data'   Ex: 2014-09-04:
    Format: :yy-mm-dd   From 'the first date of the data' to yy-mm-dd   Ex: :2025-09-04
    * Caution: Only take effect when --fetch_data is set.''')

	args = parser.parse_args()
	# import pdb; pdb.set_trace()
	cfg = {}
	if args.finmind_token is not None: cfg['finmind_token'] = args.finmind_token
	if args.enable_async is not None: cfg['enable_async'] = True
	if args.stock_symbol_list is not None: cfg['stock_symbol_string'] = args.stock_symbol_list
	if args.data_date_range is not None: cfg['data_date_range_string'] = args.data_date_range
	# import pdb; pdb.set_trace()
	with AlphaStackAnalyzer(cfg) as obj:
		if args.analyze:
			obj.analyze()
			sys.exit(0)