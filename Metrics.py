#!/usr/bin/env python
# coding: utf-8

# ## Importing Necessary Libraries

# In[1]:


import requests
import pandas as pd
import numpy as np
import os
import warnings
warnings.filterwarnings("ignore")
import time
import zipfile
from urllib.request import Request, urlopen 
import urllib.request
import collections


# In[3]:


class Stock_Metric:
    ## Function to download PDFs
    def download_url(self, url, save_path, chunk_size=128):
        r = requests.get(url, stream=True)
        with open(save_path, 'wb') as fd:
            for chunk in r.iter_content(chunk_size=chunk_size):
                fd.write(chunk)

    
    ## Compounded Annual Growth Rate for Revenue
    def Rev_CAGR(self, df, a):
        #Creating a new dataframe with only the necessary data
        new_df = df[['Sales -']]
        new_df['Sales -'] = new_df['Sales -'].str.replace(',', '')
        new_df['Sales -'] = pd.to_numeric(new_df['Sales -'])
        met = {}
        met['Margin'] = 10

        # Calculating CAGR from available data
        if len(new_df) < 10:
            met['10 years of data'] = 'Not available'
            n = len(new_df)
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[0, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            met['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 10:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Fail'
            else:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Pass'
        else:
            n = 10
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[-10, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            met['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 10:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Fail'
            else:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Pass'
        a['Revenue CAGR'] = met
    
    ## Debt to Equity Ratio
    def Debt_Eq(self, df, a):
        # Extracting necessary data from Top Ratios
        dte = df.iloc[0, 9]
        dte = float(dte)
        met = {}
        met['Margin'] = 0.5
        if np.isnan(dte) == True:
            dte = np.nan_to_num(dte)

        # Checking if Debt to Equity ratio qualifies
        met['Debt to Equity Ratio'] = dte
        if dte > 0.5:
            met['Debt to Equity Ratio metric'] = 'Fail'
        else:
            met['Debt to Equity Ratio metric'] = 'Pass'
        a['Debt to Equity Ratio'] = met

    ## Return on Assets
    def ROA(self, df, a):
        # Creating a new Dataframe with only the necessary data
        new_df = df[['Sales -', 'Expenses -', 'Total Assets']]
        new_df['Sales -'] = new_df['Sales -'].str.replace(',', '')
        new_df['Sales -'] = pd.to_numeric(new_df['Sales -'])
        new_df['Expenses -'] = new_df['Expenses -'].str.replace(',', '')
        new_df['Expenses -'] = pd.to_numeric(new_df['Expenses -'])
        new_df['Total Assets'] = new_df['Total Assets'].str.replace(',', '')
        new_df['Total Assets'] = pd.to_numeric(new_df['Total Assets'])
        new_df['Net Income'] = new_df['Sales -'] - new_df['Expenses -']
        new_df['ROA'] = (new_df['Net Income'] / new_df['Total Assets']) * 100
        new_df['ROA'] = new_df['ROA'].fillna(0)
        met = {}
        met['Margin'] = 5

        # Checking for year-wise Return on Assets
        lst = {}
        year_count = 0
        for row in new_df.itertuples():
            lst[row.Index] = row.ROA
            if row.ROA < 5:
                year_count = year_count + 1
        met['Year-wise Return on Assets'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            met['Return on Assets metric'] = 'Fail'
        else:
            met['Return on Assets metric'] = 'Pass'
        a['Return on Assets'] = met

    ## Inventory Turnover Ratio
    def Invt(self, df, a):
        # Extracting necessary data from Top Ratios
        itr = df.iloc[0, 11]
        met = {}
        met['Margin'] = [5, 10]
        if itr == '':
            itr = 0
        else:
            itr = float(itr)
        if np.isnan(itr) == True:
            itr = np.nan_to_num(itr)

        # Checking if Inventory Turnover Ratio qualifies threshold
        met['Inventory Turnover Ratio'] = itr
        if itr == 0:
            met['Inventory Turnover Ratio metric'] = 'Not enough data!'
        elif itr < 5 or itr > 10:
            met['Inventory Turnover Ratio metric'] = 'Fail'
        else:
            met['Inventory Turnover Ratio metric'] = 'Pass'
        a['Inventory Turnover Ratio'] = met

    ## Cash Conversion Cycle
    def CCC(self, df, df1, a):
        # Extracting necessary data from Top Ratios and Market Leader
        met = {}
        ccc = df.iloc[0, 12]
        ccc = str(ccc)
        ccc = ccc.replace(',', '')
        ccc = float(ccc)
        ml = df1.iloc[0, 12]
        ml = str(ml)
        ml = ml.replace(',', '')
        ml = float(ml)
        if np.isnan(ccc) == True:
            ccc = np.nan_to_num(ccc)
        if np.isnan(ml) == True:
            ml = np.nan_to_num(ml)
        met['Margin'] = 'Cash Conversion Cycle of Market Leader'

        # Comparing Cash Conversion Cycles
        met['Cash Conversion Cycle of Company'] = ccc
        met['Cash Conversion Cycle of Market Leader'] = ml
        
        if ccc > ml:
            met['Cash Conversion Cycle'] = 'Fail'
        else:
            met['Cash Conversion Cycle'] = 'Pass'
        a['Cash Conversion Cycle'] = met

    ## Earnings per Share
    def EPS(self, df, a):
        #Creating a new dataframe with the necessary data
        new_df = df[['EPS in Rs']]
        new_df['EPS in Rs'] = new_df['EPS in Rs'].str.replace(',', '')
        try:
            new_df['EPS in Rs'] = pd.to_numeric(new_df['EPS in Rs'])
        except:
            new_df['EPS in Rs'] = new_df['EPS in Rs'].astype(float)
        new_df.fillna(0, inplace = True)
        new_df['Difference in EPS'] = new_df['EPS in Rs'] - new_df['EPS in Rs'].shift(1)
        new_df.fillna(0, inplace = True)
        new_df['EPS_Growth_Rate'] = (new_df['Difference in EPS'] / new_df['EPS in Rs'].shift(1)) * 100
        new_df.fillna(0, inplace = True)
        met = {}
        met['Margin'] = 25

        #Checking for year-wise Earnings per Share
        year_count = 0
        lst = {}
        for row in new_df.itertuples():
            lst[row.Index] = row.EPS_Growth_Rate
            if row.EPS_Growth_Rate < 25:
                year_count = year_count + 1
        met['Year-wise EPS Growth Rate'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            met['Earnings per Share metric'] = 'Fail'
        else:
            met['Earnings per Share metric'] = 'Pass'
        a['Earnings per Share'] = met

    ## Operating Cash Flow
    def OCF(self, df, df1, a):
        ocf = df.iloc[0, 13]
        ocf = ocf.replace('₹ ', '')
        ocf = ocf.replace(' Cr.', '')
        ocf = ocf.replace(',', '')
        ocf = float(ocf)
        if np.isnan(ocf) == True:
            ocf = np.nan_to_num(ocf)
        met = {}
        met['Margin'] = 'Net Income for last 3 years'
        met['Operating Cash Flow for last 3 years'] = ocf

        new_df = df1[['Sales -', 'Expenses -']]
        new_df['Sales -'] = new_df['Sales -'].str.replace(',', '')
        new_df['Sales -'] = pd.to_numeric(new_df['Sales -'])
        new_df['Expenses -'] = new_df['Expenses -'].str.replace(',', '')
        new_df['Expenses -'] = pd.to_numeric(new_df['Expenses -'])
        new_df['Net Income'] = new_df['Sales -'] - new_df['Expenses -']
        new_df['Net Income'] = new_df['Net Income'].fillna(0)

        last_three = new_df['Net Income'].tail(3)
        net_incm = last_three.mean()

        met['Net Income for last 3 years'] = net_incm
        
        if ocf < net_incm:
            met['Operating Cash Flow'] = 'Fail'
        else:
            met['Operating Cash Flow'] = 'Pass'
        a['Operating Cash Flow'] = met

    ## CAGR for ROCE
    def ROCE(self, df, a):
        new_df = df[['ROCE %']]
        new_df['ROCE %'] = new_df['ROCE %'].str.replace('%', '')
        new_df['ROCE %'] = pd.to_numeric(new_df['ROCE %'])
        new_df.fillna(0, inplace=True)
        met = {}
        met['Margin'] = 15

        if len(new_df) < 10:
            met['10 years of data'] = 'Not available'
            ev = new_df.iloc[-1, 0]
            if new_df.iloc[0, 0] == 0:
                n = len(new_df) - 1
                bv = new_df.iloc[1, 0]
            else:
                n = len(new_df)
                bv = new_df.iloc[0, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            met['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 15:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Fail'
            else:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Pass'
        else:
            n = 10
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[-10, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            met['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 15:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Fail'
            else:
                met['Compounded Annual Growth Rate for Revenue metric'] = 'Pass'
        a['ROCE'] = met

    ## Free Cash Flow
    def FCF(self, df, a):
        new_df = df[['Cash from Operating Activity -', 'Cash from Investing Activity -']]
        new_df['Cash from Operating Activity -'] = new_df['Cash from Operating Activity -'].str.replace(',', '')
        new_df['Cash from Operating Activity -'] = pd.to_numeric(new_df['Cash from Operating Activity -'])
        new_df['Cash from Investing Activity -'] = new_df['Cash from Investing Activity -'].str.replace(',', '')
        new_df['Cash from Investing Activity -'] = pd.to_numeric(new_df['Cash from Investing Activity -'])
        new_df['FCF'] = new_df['Cash from Operating Activity -'] + new_df['Cash from Investing Activity -']
        new_df['FCF'] = new_df['FCF'].fillna(0)
        met = {}
        met['Margin'] = 0

        year_count = 0
        lst = {}
        for row in new_df.itertuples():
            lst[row.Index] = row.FCF
            if row.FCF < 0:
                year_count = year_count + 1
        met['Year-wise Free Cash Flow'] = lst
        if year_count > 0:
            met['Free Cash Flow metric'] = 'Fail'
        else:
            met['Free Cash Flow metric'] = 'Pass'
        a['Free Cash Flow'] = met

    ## Interest Coverage Ratio
    def Intrst(self, df, a):
        icr = df.iloc[0, 10]
        met = {}
        met['Margin'] = [24, 100]
        if icr == '':
            icr = 0
        else:
            icr = float(icr)
        if np.isnan(icr) == True:
            icr = np.nan_to_num(icr)

        met['Interest Coverage Ratio'] = icr
        if icr < 24 or icr > 100:
            met['Interest Coverage Ratio metric'] = 'Fail'
        else:
            met['Interest Coverage Ratio metric'] = 'Pass'
        a['Interest Coverage Ratio'] = met

    ## CFO as % of EBITDA
    def CFO(self, df, a):
        new_df = df[['Cash from Financing Activity -', 'Operating Profit', 'Other Income -']]
        new_df['Cash from Financing Activity -'] = new_df['Cash from Financing Activity -'].str.replace(',', '')
        new_df['Operating Profit'] = new_df['Operating Profit'].str.replace(',', '')
        new_df['Other Income -'] = new_df['Other Income -'].str.replace(',', '')
        new_df['Cash from Financing Activity -'] = pd.to_numeric(new_df['Cash from Financing Activity -'])
        new_df['Operating Profit'] = pd.to_numeric(new_df['Operating Profit'])
        new_df['Other Income -'] = pd.to_numeric(new_df['Other Income -'])
        new_df['EBITDA'] = new_df['Operating Profit'] + new_df['Other Income -']
        new_df['CFO %'] = (new_df['Cash from Financing Activity -'] / new_df['EBITDA']) * 100

        mean = new_df['CFO %'].mean()
        std = new_df['CFO %'].std()
        new_df['CFO_Zscore'] = (new_df['CFO %'] - mean) / std
        new_df['CFO_Zscore'] = new_df['CFO_Zscore'].fillna(0)
        met = {}
        met['Margin'] = [-1, 1]

        year_count = 0
        lst = {}
        for row in new_df.itertuples():
            lst[row.Index] = row.CFO_Zscore
            if row.CFO_Zscore < -1 or row.CFO_Zscore > 1:
                year_count = year_count + 1
        met['Year-wise CFO Z-Score'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            met['CFO as % of EBITDA metric'] = 'Fail'
        else:
            met['CFO as % of EBITDA metric'] = 'Pass'
        a['CFO'] = met


    ## Depreciation Rates
    def Depr(self, df, a):
        new_df = df[['Depreciation']]
        new_df['Depreciation'] = new_df['Depreciation'].str.replace(',', '')
        new_df['Depreciation'] = pd.to_numeric(new_df['Depreciation'])
        new_df['Difference in Depreciation'] = new_df['Depreciation'] - new_df['Depreciation'].shift(1)
        new_df.fillna(0, inplace=True)
        new_df['Depreciation Growth Rate'] = (new_df['Difference in Depreciation'] / new_df['Depreciation'].shift(
            1)) * 100
        new_df.fillna(0, inplace=True)

        mean = new_df['Depreciation Growth Rate'].mean()
        std = new_df['Depreciation Growth Rate'].std()
        new_df['Depreciation_Zscore'] = (new_df['Depreciation Growth Rate'] - mean) / std
        new_df['Depreciation_Zscore'] = new_df['Depreciation_Zscore'].fillna(0)
        met = {}
        met['Margin'] = [-2, 2]

        year_count = 0
        lst = {}
        for row in new_df.itertuples():
            lst[row.Index] = row.Depreciation_Zscore
            if row.Depreciation_Zscore < -2 or row.Depreciation_Zscore > 2:
                year_count = year_count + 1
        met['Year-wise Depreciation Z-Score'] = lst
        if year_count > 0:
            met['Changes in Depreciation metric'] = 'Fail'
        else:
            met['Changes in Depreciation metric'] = 'Pass'
        a['Depreciation Rates'] = met


    ## Change in Reserves
    def Rsrv(self, df, a):
        new_df = df[['Reserves']]
        new_df['Reserves'] = new_df['Reserves'].str.replace(',', '')
        new_df['Reserves'] = pd.to_numeric(new_df['Reserves'])
        new_df['Difference'] = new_df['Reserves'] - new_df['Reserves'].shift(1)
        new_df.fillna(0, inplace=True)

        mean = new_df['Difference'].mean()
        std = new_df['Difference'].std()
        new_df['Reserves_Zscore'] = (new_df['Difference'] - mean) / std
        met = {}
        met['Margin'] = [-2, 2]

        year_count = 0
        lst = {}
        for row in new_df.itertuples():
            lst[row.Index] = row.Reserves_Zscore
            if row.Reserves_Zscore < -2 or row.Reserves_Zscore > 2:
                year_count = year_count + 1
        met['Year-wise Reserves Z-Score'] = lst
        if year_count > 0:
            met['Changes in Reserve metric'] = 'Fail'
        else:
            met['Changes in Reserve metric'] = 'Pass'
        a['Changes in Reserves'] = met


    ## Yields on Cash and Cash Equivalents
    def Cash(self, df, a):
        new_df = df[['Interest']]
        new_df['Interest'] = new_df['Interest'].str.replace(',', '')
        new_df['Interest'] = pd.to_numeric(new_df['Interest'])
        new_df['Difference'] = new_df['Interest'] - new_df['Interest'].shift(1)
        new_df.fillna(0, inplace=True)

        mean = new_df['Difference'].mean()
        std = new_df['Difference'].std()
        new_df['Interest_Zscore'] = (new_df['Difference'] - mean) / std
        met = {}
        met['Margin'] = [-1, 1]

        year_count = 0
        lst = {}
        for row in new_df.itertuples():
            lst[row.Index] = row.Interest_Zscore
            if row.Interest_Zscore < -1 or row.Interest_Zscore > 1:
                year_count = year_count + 1
        met['Year-wise Interest Z-Score'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            met['Changes in Interest metric'] = 'Fail'
        else:
            met['Changes in Interest metric'] = 'Pass'
        a['Yields on Cash'] = met


    ## Contingent Liabilities as % of Net Worth
    def Cont(self, df, a):
        cont = df.iloc[0, -2]
        worth = df.iloc[0, -1]
        cont = cont.replace('₹', '')
        worth = worth.replace('₹', '')
        cont = cont.replace('Cr.', '')
        worth = worth.replace('Cr.', '')
        cont = cont.replace(',', '')
        worth = worth.replace(',', '')
        cont = float(cont)
        worth = int(worth)

        liab = (cont / worth) * 100
        met = {}
        met['Margin'] = 25

        if liab > 25:
            met['Contingent Liability metric'] = 'Fail'
        else:
            met['Contingent Liability metric'] = 'Pass'
        a['Contingent Liabilities'] = met

    ## CWIP to Gross Block
    def CWIP(self, df, a):
        new_df = df[['CWIP', 'Gross Block']]
        new_df['CWIP'] = new_df['CWIP'].str.replace(',', '')
        new_df['CWIP'] = pd.to_numeric(new_df['CWIP'])
        new_df['Gross Block'] = new_df['Gross Block'].str.replace(',', '')
        new_df['Gross Block'] = pd.to_numeric(new_df['Gross Block'])
        new_df['Ratios'] = new_df['CWIP'] / new_df['Gross Block']
        new_df['Ratios'] = new_df['Ratios'].fillna(0)
        met = {}
        met['Margin'] = [0.05, 0.75]

        year_count = 0
        lst = {}
        for row in new_df.itertuples():
            lst[row.Index] = row.Ratios
            if row.Ratios < 0.05:
                year_count = year_count + 1
            elif row.Ratios > 0.75:
                year_count = year_count + 1
        met['Year-wise CWIP TO Gross Block Ratios'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            met['CWIP to Gross Block metric'] = 'Fail'
        else:
            met['CWIP to Gross Block metric'] = 'Pass'
        a['CWIP'] = met


    def screener(self, lst):
        sym = lst[0]
        print(os.path.abspath(sym))
        path1 = os.path.abspath(sym)

        df_final = pd.read_excel(path1 + '\\' + 'balance_sheet_cas.xlsx')
        peer_comparison = pd.read_excel(path1 + '\\' + 'peer_comparison.xlsx')
        top_ratios = pd.read_excel(path1 + '\\' + 'top_ratios.xlsx')
        market_leader = pd.read_excel(path1 + '\\' + 'market_leader.xlsx')

        transpose_df = df_final.transpose()
        transpose_df = transpose_df.drop(transpose_df.index[-1])
        transpose_df = transpose_df.set_axis(transpose_df.iloc[0], axis = 1).drop(transpose_df.index[0])
        
        a ={}
        self.ROA(transpose_df, a)
        self.EPS(transpose_df, a)
        self.FCF(transpose_df, a)
        self.CFO(transpose_df, a)
        self.Depr(transpose_df, a)
        self.Rsrv(transpose_df, a)
        self.Cash(transpose_df, a)
        self.CWIP(transpose_df, a)
        
        self.Rev_CAGR(transpose_df, a)
        self.Debt_Eq(top_ratios, a)
        self.Invt(top_ratios, a)
        self.CCC(top_ratios, market_leader, a)
        self.OCF(top_ratios, transpose_df, a)
        self.ROCE(transpose_df, a)
        self.Intrst(top_ratios, a)
        self.Cont(top_ratios, a)
        
        return a
