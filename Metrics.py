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

        # Calculating CAGR from available data
        if len(new_df) < 10:
            a['10 years of data'] = 'Not available'
            n = len(new_df)
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[0, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            a['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 10:
                a['Compounded Annual Growth Rate for Revenue metric'] = 'Failing!'
        else:
            n = 10
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[-10, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            a['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 10:
                a['Compounded Annual Growth Rate for Revenue metric'] = 'Failing!'
    
    ## Debt to Equity Ratio
    def Debt_Eq(self, df, a):
        # Extracting necessary data from Top Ratios
        dte = df.iloc[0, 9]
        dte = float(dte)

        # Checking if Debt to Equity ratio qualifies
        a['Debt to Equity Ratio'] = dte
        if dte > 0.5:
            a['Debt to Equity Ratio metric'] = 'Failing!'

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

        # Checking for year-wise Return on Assets
        lst = []
        year_count = 0
        for row in new_df.itertuples():
            if row.ROA < 5:
                year_count = year_count + 1
                lst.append({row.Index})
        a['Years with Alarming Returns on Assets'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            a['Return on Assets metric'] = 'Failing!'

    ## Inventory Turnover Ratio
    def Invt(self, df, a):
        # Extracting necessary data from Top Ratios
        itr = df.iloc[0, 11]
        if itr == '':
            itr = 0
        else:
            itr = float(itr)

        # Checking if Inventory Turnover Ratio qualifies threshold
        a['Inventory Turnover Ratio'] = itr
        if itr == 0:
            a['Inventory Turnover Ratio metric'] = 'Not enough data!'
        elif itr < 5 or itr > 10:
            a['Inventory Turnover Ratio metric'] = 'Failing!'

    ## Cash Conversion Cycle
    def CCC(self, df, df1, a):
        # Extracting necessary data from Top Ratios and Market Leader
        ccc = df.iloc[0, 12]
        ccc = str(ccc)
        ccc = ccc.replace(',', '')
        ccc = float(ccc)
        ml = df1.iloc[0, 12]
        ml = str(ml)
        ml = ml.replace(',', '')
        ml = float(ml)

        # Comparing Cash Conversion Cycles
        a['Cash Conversion Cycle of Company'] = ccc
        a['Cash Conversion Cycle of Market Leader'] = ml
        
        if ccc > ml:
            a['Cash Conversion Cycle'] = 'Failing!'

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

        #Checking for year-wise Earnings per Share
        year_count = 0
        lst = []
        for row in new_df.itertuples():
            if row.EPS_Growth_Rate < 25:
                year_count = year_count + 1
                lst.append({row.Index})
        a['Years with alarming EPS Growth Rate'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            a['Earnings per Share metric'] = 'Failing!'

    ## Operating Cash Flow
    def OCF(self, df, df1, a):
        ocf = df.iloc[0, 13]
        ocf = ocf.replace('₹ ', '')
        ocf = ocf.replace(' Cr.', '')
        ocf = ocf.replace(',', '')
        ocf = float(ocf)
        a['Operating Cash Flow for last 3 years'] = ocf

        new_df = df1[['Sales -', 'Expenses -']]
        new_df['Sales -'] = new_df['Sales -'].str.replace(',', '')
        new_df['Sales -'] = pd.to_numeric(new_df['Sales -'])
        new_df['Expenses -'] = new_df['Expenses -'].str.replace(',', '')
        new_df['Expenses -'] = pd.to_numeric(new_df['Expenses -'])
        new_df['Net Income'] = new_df['Sales -'] - new_df['Expenses -']

        last_three = new_df['Net Income'].tail(3)
        net_incm = last_three.mean()

        a['Net Income for last 3 years'] = net_incm
        
        if ocf < net_incm:
            a['Operating Cash Flow'] = 'Failing!'

    ## CAGR for ROCE
    def ROCE(self, df, a):
        new_df = df[['ROCE %']]
        new_df['ROCE %'] = new_df['ROCE %'].str.replace('%', '')
        new_df['ROCE %'] = pd.to_numeric(new_df['ROCE %'])
        new_df.fillna(0, inplace=True)

        if len(new_df) < 10:
            a['10 years of data'] = 'Not available'
            ev = new_df.iloc[-1, 0]
            if new_df.iloc[0, 0] == 0:
                n = len(new_df) - 1
                bv = new_df.iloc[1, 0]
            else:
                n = len(new_df)
                bv = new_df.iloc[0, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            a['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 15:
                a['Compounded Annual Growth Rate for Revenue metric'] = 'Failing!'
        else:
            n = 10
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[-10, 0]
            cagr = (pow((ev / bv), (1 / n)) - 1) * 100
            a['Compounded Annual Growth Rate for Revenue'] = cagr
            if cagr < 15:
                a['Compounded Annual Growth Rate for Revenue metric'] = 'Failing!'

    ## Free Cash Flow
    def FCF(self, df, a):
        new_df = df[['Cash from Operating Activity -', 'Cash from Investing Activity -']]
        new_df['Cash from Operating Activity -'] = new_df['Cash from Operating Activity -'].str.replace(',', '')
        new_df['Cash from Operating Activity -'] = pd.to_numeric(new_df['Cash from Operating Activity -'])
        new_df['Cash from Investing Activity -'] = new_df['Cash from Investing Activity -'].str.replace(',', '')
        new_df['Cash from Investing Activity -'] = pd.to_numeric(new_df['Cash from Investing Activity -'])
        new_df['FCF'] = new_df['Cash from Operating Activity -'] + new_df['Cash from Investing Activity -']

        year_count = 0
        lst = []
        for row in new_df.itertuples():
            if row.FCF < 0:
                year_count = year_count + 1
                lst.append(row.Index)
        a['Years with negative Free Cash Flow'] = lst
        if year_count > 0:
            a['Free Cash Flow metric'] = 'Failing!'

    ## Interest Coverage Ratio
    def Intrst(self, df, a):
        icr = df.iloc[0, 10]
        if icr == '':
            icr = 0
        else:
            icr = float(icr)

        a['Interest Coverage Ratio'] = icr
        if icr < 24 or icr > 100:
            a['Interest Coverage Ratio metric'] = 'Failing!'

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

        year_count = 0
        lst = []
        for row in new_df.itertuples():
            if row.CFO_Zscore < -1 or row.CFO_Zscore > 1:
                year_count = year_count + 1
                lst.append(row.Index)
        a['Year with Alarming CFO%'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            a['CFO as % of EBITDA metric'] = 'Failing!'


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

        year_count = 0
        lst = []
        for row in new_df.itertuples():
            if row.Depreciation_Zscore < -2 or row.Depreciation_Zscore > 2:
                year_count = year_count + 1
                lst.append(row.Index)
        a['Year with Alarming Changes in Depreciation %'] = lst
        if year_count > 0:
            a['Changes in Depreciation metric'] = 'Failing!'


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

        year_count = 0
        lst = []
        for row in new_df.itertuples():
            if row.Reserves_Zscore < -2 or row.Reserves_Zscore > 2:
                year_count = year_count + 1
                lst.append(row.Index)
        a['Year with Alarming Changes in Reserves'] = lst
        if year_count > 0:
            a['Changes in Reserve metric'] = 'Failing!'


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

        year_count = 0
        lst = []
        for row in new_df.itertuples():
            if row.Interest_Zscore < -1 or row.Interest_Zscore > 1:
                year_count = year_count + 1
                lst.append(row.Index)
        a['Year with Alarming Changes in Interest'] = lst
        n = int(0.75 * len(new_df))
        if year_count > n:
            a['Changes in Interest metric'] = 'Failing!'
        return a


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

        if liab > 25:
            a['Contingent Liability metric'] = 'Failing!'

    ## CWIP to Gross Block
    def CWIP(self, df, a):
        new_df = df[['CWIP', 'Gross Block']]
        new_df['CWIP'] = new_df['CWIP'].str.replace(',', '')
        new_df['CWIP'] = pd.to_numeric(new_df['CWIP'])
        new_df['Gross Block'] = new_df['Gross Block'].str.replace(',', '')
        new_df['Gross Block'] = pd.to_numeric(new_df['Gross Block'])
        new_df['Ratios'] = new_df['CWIP'] / new_df['Gross Block']

        year_count = 0
        lst1 = []
        lst2 = []
        for row in new_df.itertuples():
            if row.Ratios < 0.05:
                year_count = year_count + 1
                lst1.append(row.Index)
            elif row.Ratios > 0.75:
                year_count = year_count + 1
                lst2.append(row.Index)
        a['Year with Alarming Decrease in CWIP to Gross Block Ratios'] = lst1
        a['Year with Alarming Increase in CWIP to Gross Block Ratios'] = lst2
        n = int(0.75 * len(new_df))
        if year_count > n:
            a['CWIP to Gross Block metric'] = 'Failing!'


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
        self.Invt(top_ratios, a)
        self.CCC(top_ratios, market_leader, a)
        self.EPS(transpose_df, a)
        self.OCF(top_ratios, transpose_df, a)

        self.Rev_CAGR(transpose_df, a)
        self.Debt_Eq(top_ratios, a)
        self.ROCE(transpose_df, a)
        self.FCF(transpose_df, a)
        self.Intrst(top_ratios, a)
        self.CFO(transpose_df, a)
        self.Depr(transpose_df, a)
        self.Rsrv(transpose_df, a)
        self.Cash(transpose_df, a)
        self.Cont(top_ratios, a)
        self.CWIP(transpose_df, a)
        
        return a
