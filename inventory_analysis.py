import numpy as np
import pandas as pd
from datetime import datetime
from pandas import ExcelWriter
from os import listdir
from os.path import isfile, join
import re
from time import strptime
import random
from openpyxl.utils import range_boundaries
from collections import defaultdict, Counter
import itertools
import sys


class InventoryRates():
    def __init__(self, inv_type='cumulative', data_path = None, data_sheet = 0, date_col = None, issue_col = None,soh_col = None, window = 12, agg_cols = [],fac_cols=[],prod_cols=[]):
        self.inv_type = inv_type
        self.data_path = data_path
        self.data_sheet = data_sheet
        self.data_frame = pd.read_excel(self.data_path,sheet_name = self.data_sheet)
        self.data_frame = self.data_frame.rename(columns=lambda x: x.strip())
        self.date_col = date_col
        self.issue_col = issue_col
        self.soh_col = soh_col
        self.window = window
        self.agg_cols = agg_cols
        self.rates_data = pd.DataFrame()
        self.stock_data = pd.DataFrame()
        self.fac_cols = fac_cols
        self.prod_cols = prod_cols
        self.fac_table = pd.DataFrame(self.data_frame[self.fac_cols[0]].unique())
        self.prod_table = pd.DataFrame(self.data_frame[self.prod_cols[0]].unique())
    
    @property
    def inv_type(self):
        return self._inv_type

    @inv_type.setter
    def inv_type(self, t):
        if not (t in ['transactional', 'cumulative']): 
            raise ValueError("Invalid inventory data type")
        self._inv_type = t

    @property
    def data_frame(self):
    	return self._data_frame
    
    @data_frame.setter
    def data_frame(self, df):
    	self._data_frame = df

    @property
    def date_col(self):
    	return self._date_col

    @date_col.setter
    def date_col(self, d):
    	if not d:
    		raise ValueError(f"Date column not specified")
    	if not (d in self.data_frame):
    		raise ValueError(f"Column '{d}' is not in data")
    	self._date_col = d
    	self.data_frame[self.date_col] = pd.to_datetime(self.data_frame[self.date_col])
    	# if cumulative, check that date does not occur multiple times for the same grouping

    @property
    def issue_col(self):
    	return self._issue_col

    @issue_col.setter
    def issue_col(self, i):
    	if not i:
    		raise ValueError(f"Issue column not specified")
    	if not (i in self.data_frame):
    		raise ValueError(f"Column '{i}' is not in data")
    	self._issue_col = i

    @property
    def soh_col(self):
    	return self._soh_col

    @soh_col.setter
    def soh_col(self, s):
    	if not s:
    		raise ValueError(f"SOH column not specified")
    	if not (s in self.data_frame):
    		raise ValueError(f"Column '{s}' is not in data")
    	self._soh_col = s

    @property
    def agg_cols(self):
    	return self._agg_cols

    @agg_cols.setter
    def agg_cols(self, a):
    	if not a:
    		raise ValueError(f"Aggregation columns not specified")
    	for col in a:
    		if not (col in self.data_frame):
    			raise ValueError(f"Column '{col}' is not in data")
    	self._agg_cols = a

    @property
    def fac_cols(self):
    	return self._fac_cols

    @fac_cols.setter
    def fac_cols(self, f):
    	if not f:
    		raise ValueError(f"Facility columns not specified")
    	for col in f:
    		if not (col in self.data_frame):
    			raise ValueError(f"Column '{col}' is not in data")
    	self._fac_cols = f

    @property
    def prod_cols(self):
    	return self._prod_cols

    @prod_cols.setter
    def prod_cols(self, p):
    	if not p:
    		raise ValueError(f"Product columns not specified")
    	for col in p:
    		if not (col in self.data_frame):
    			raise ValueError(f"Column '{col}' is not in data")
    	self._prod_cols = p

    @property
    def window(self):
    	return self._window

    @window.setter
    def window(self, w):
    	#if not (len(w) == 2):
    	#	raise ValueError("Invalid window - must have length 2 min and max")
    	if not (type(w) == int): #and type(w[1] == int)):
    		raise ValueError("Invalid window types - must be integer")
    	#if not (w[0] < w[1]):
    		#raise ValueError("Invalid window - min must be less than max")
    	self._window = w

    #@method
    def rolling_rates(self):
    	for col in self.agg_cols:
    		self.data_frame[col] = self.data_frame[col].map(str)
    	grouped = self.data_frame.groupby(self.agg_cols+[self.date_col]).agg({self.issue_col:'sum', self.soh_col:'mean'}).groupby(level = 0).rolling(window=self.window, min_periods=self.window).agg({self.issue_col:['std','mean','sum','count'],self.soh_col:'mean'}).reset_index(level=0,drop=True).reset_index()
    	grouped.columns = grouped.columns.map('_'.join).str.strip('_')
    	grouped['Consumption_COV'] = grouped[self.issue_col+'_std']/grouped[self.issue_col+'_mean']
    	grouped['Inventory_turn'] = grouped[self.issue_col+'_sum']/grouped[self.soh_col+'_mean']
    	self.rates_data = grouped

    def stock_status(self):
    	return None




    
    
    