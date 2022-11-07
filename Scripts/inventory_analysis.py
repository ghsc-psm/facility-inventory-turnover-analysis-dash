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
        print(">> reading in data")
        self.data_frame = pd.read_excel(self.data_path,sheet_name = self.data_sheet) if ('.xlsx' in self.data_path) else pd.read_csv(self.data_path)
        self.data_frame = self.data_frame.rename(columns=lambda x: x.strip())
        print(">> data read in ")
        print(">> normalizing date types")
        self.date_col = date_col
        self.issue_col = issue_col
        self.soh_col = soh_col
        self.window = window
        self.agg_cols = agg_cols
        self.rates_data = pd.DataFrame()
        self.stock_data = pd.DataFrame()
        self.fac_cols = fac_cols
        self.prod_cols = prod_cols
    
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
        def groupby_fn(x):
            # calculate rolling rates
            df = x.rolling(window = self.window, min_periods = self.window//2).agg({self.issue_col:['std','mean','sum','count'],self.soh_col:'mean'})
           
            # collapse column names
            df.columns = df.columns.map('_'.join).str.strip('_')

            # combine with grouped columns
            df = pd.concat([x[self.agg_cols+[self.date_col]],df],axis=1,join='inner')

            return df

        for col in self.agg_cols:
            self.data_frame[col] = self.data_frame[col].map(str)
        print('>> aggregating monthly values')
        grouped = (self.data_frame.groupby(self.agg_cols+[self.date_col])
                   .agg({self.issue_col:'sum', self.soh_col:'mean'})).reset_index()

        print(">> calculating rolling rates by facility/product (this may take a moment)")
        grouped = grouped.groupby(self.agg_cols).apply(lambda x: groupby_fn(x))
        
        grouped['Consumption_COV'] = grouped[self.issue_col+'_std']/grouped[self.issue_col+'_mean']
        grouped['Inventory_turn'] = grouped[self.issue_col+'_sum']/grouped[self.soh_col+'_mean']
        grouped.loc[~np.isfinite(grouped['Inventory_turn']), 'Inventory_turn'] = np.nan
        grouped.loc[~np.isfinite(grouped['Consumption_COV']),'Consumption_COV'] = np.nan
        self.rates_data = grouped

    #@method
    def stock_status(self):
        
        def get_blank(x):
            return (pd.isnull(x[self.issue_col])) & (pd.isnull(x[self.soh_col]))
        def group_last_soh(x):
            d = {}
            x['Blank'] = x.apply(get_blank,axis=1)
            #d['Group'] = x[unique_group].sum(1)
            x_non_blank = x[x['Blank'] == False]
            if x_non_blank.shape[0] == 0:
                d['Stock Status As Of'] = np.nan
                d['Stock on Hand'] = np.nan
            else:
                d['Stock Status As Of'] = x_non_blank[self.date_col].iloc[-1]
                d['Stock on Hand'] = np.nan if pd.isnull(x_non_blank[self.soh_col].iloc[-1]) else x_non_blank[self.soh_col].iloc[-1] 
            x_cons = x[x[self.issue_col]>0]
            if x_cons.shape[0] < 3:
                d['AMC'] = x_cons[self.issue_col].mean()
            else:
                d['AMC'] = x_cons.iloc[[-3,-2,-1]][self.issue_col].mean()
            d['% Records Blank'] = (x[x['Blank'] == True].shape[0])/x.shape[0]
            return pd.Series(d,index=['Stock Status As Of','Stock on Hand','AMC','% Records Blank'])

        df = self.data_frame.sort_values(by=[self.date_col])
        df = df.groupby([self.fac_cols[0],self.prod_cols[0]]).apply(lambda x: group_last_soh(x)).reset_index()
        df['MOS'] = np.where(df['AMC']==0,0,df['Stock on Hand']/df['AMC'])
        df = df[[self.fac_cols[0],self.prod_cols[0],'Stock Status As Of','Stock on Hand','AMC','MOS','% Records Blank']]
        self.stock_data = df

    def transactional(self):
        if self.inv_type == 'transactional':
            self.data_frame[self.date_col] = self.data_frame[self.date_col].apply(lambda x: x.replace(day=1))

    def format_columns(self):
        names = {self.fac_cols[0]:'Facility_id',
             self.prod_cols[0]:'Product_id',
             self.date_col:'Date',
             self.issue_col+'_std':'Consumption_std',
             self.issue_col+'_mean':'Consumption_mean',
             self.issue_col+'_sum':'Consumption_sum',
             self.issue_col+'_count':'Consumption_count',
             self.soh_col+'_mean':'SOH_mean'}
        self.rates_data = self.rates_data.rename(columns=names)
        self.rates_data = self.rates_data[list(names.values())+['Consumption_COV','Inventory_turn']]
        self.stock_data = self.stock_data.rename(columns={self.fac_cols[0]:'Facility_id',self.prod_cols[0]:'Product_id'})

    def clean_text(self):
        for col1 in self.prod_cols:
            self.data_frame[col1] = self.data_frame[col1].apply(lambda x: str(x).strip().upper())  
        for col2 in self.fac_cols:
            self.data_frame[col2] = self.data_frame[col2].apply(lambda x: str(x).strip().upper())

    def create_ref_tables(self):
        self.fac_table = self.data_frame[self.fac_cols].drop_duplicates(subset=self.fac_cols[0])
        self.prod_table = self.data_frame[self.prod_cols].drop_duplicates(subset=self.prod_cols[0]).dropna()
        self.plot_data = self.rates_data[['Facility_id','Product_id']].drop_duplicates()
        self.dates_table = pd.DataFrame(sorted(list(self.rates_data['Date'].unique())),columns=['Date'])





    
    
    