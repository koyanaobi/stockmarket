#!/usr/bin/env python
# coding: utf-8

# In[ ]:


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
import glob


# In[2]:


class Stock_Market:
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

        #Calculating CAGR from available data
        if len(new_df) < 10:
            n = len(new_df)
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[0, 0]
            cagr = (pow((ev / bv), (1/n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            if cagr < 10:
                a = a + 1
        else:
            n = 10
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[-10, 0]
            cagr = (pow((ev / bv), (1/n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            if cagr < 10:
                a = a + 1
        return a
    
    ## Debt to Equity Ratio
    def Debt_Eq(self, df, a):
        #Extracting necessary data from Top Ratios
        dte = df.iloc[0, 9]
        dte = float(dte)
        if np.isnan(dte) == True:
            dte = np.nan_to_num(dte)

        #Checking if Debt to Equity ratio qualifies
        if dte > 0.5:
            a = a + 1
        return a
    
    ## Return on Assets
    def ROA(self, df, a):
        #Creating a new Dataframe with only the necessary data
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

        #Checking for year-wise Return on Assets
        year_count = 0
        for row in new_df.itertuples():
            if row.ROA < 5:
                year_count = year_count + 1
        n = int(0.75 * len(new_df))
        if year_count > n:
            a = a + 1
        return a
    
    ## Inventory Turnover Ratio
    def Invt(self, df, a):
        #Extracting necessary data from Top Ratios
        itr = df.iloc[0, 11]
        if itr == '':
            itr = 0
        else:
            itr = float(itr)
        if np.isnan(itr) == True:
            itr = np.nan_to_num(itr)

        #Checking if Inventory Turnover Ratio qualifies threshold
        if itr < 5 or itr > 10:
            a = a + 1
        return a
    
    ## Cash Conversion Cycle
    def CCC(self, df, df1, a):
        #Extracting necessary data from Top Ratios and Market Leader
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

        #Comparing Cash Conversion Cycles
        if ccc > ml:
            a = a + 1
        return a
    
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
        for row in new_df.itertuples():
            if row.EPS_Growth_Rate < 25:
                year_count = year_count + 1
        n = int(0.75 * len(new_df))
        if year_count > n:
            a = a + 1
        return a
    
    ## Operating Cash Flow
    def OCF(self, df, df1, a):
        ocf = df.iloc[0, 13]
        ocf = ocf.replace('₹ ', '')
        ocf = ocf.replace(' Cr.', '')
        ocf = ocf.replace(',', '')
        ocf = float(ocf)
        if np.isnan(ocf) == True:
            ocf = np.nan_to_num(ocf)

        new_df = df1[['Sales -', 'Expenses -']]
        new_df['Sales -'] = new_df['Sales -'].str.replace(',', '')
        new_df['Sales -'] = pd.to_numeric(new_df['Sales -'])
        new_df['Expenses -'] = new_df['Expenses -'].str.replace(',', '')
        new_df['Expenses -'] = pd.to_numeric(new_df['Expenses -'])
        new_df['Net Income'] = new_df['Sales -'] - new_df['Expenses -']
        new_df['Net Income'] = new_df['Net Income'].fillna(0)

        last_three = new_df['Net Income'].tail(3)
        net_incm = last_three.mean()

        if ocf < net_incm:
            a = a + 1

        return a
    
    ## CAGR for ROCE
    def ROCE(self, df, a):
        new_df = df[['ROCE %']]
        new_df['ROCE %'] = new_df['ROCE %'].str.replace('%', '')
        new_df['ROCE %'] = pd.to_numeric(new_df['ROCE %'])
        new_df.fillna(0, inplace = True)

        if len(new_df) < 10:
            ev = new_df.iloc[-1, 0]
            if new_df.iloc[0, 0] == 0:
                n = len(new_df) - 1
                bv = new_df.iloc[1, 0]
            else:
                n = len(new_df)
                bv = new_df.iloc[0, 0]
            cagr = (pow((ev / bv), (1/n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            if cagr < 15:
                a = a + 1
        else:
            n = 10
            ev = new_df.iloc[-1, 0]
            bv = new_df.iloc[-10, 0]
            cagr = (pow((ev / bv), (1/n)) - 1) * 100
            if np.isnan(cagr) == True:
                cagr = np.nan_to_num(cagr)
            if cagr < 15:
                a = a + 1
        return a
    
    ## Free Cash Flow
    def FCF(self, df, a):
        new_df = df[['Cash from Operating Activity -', 'Cash from Investing Activity -']]
        new_df['Cash from Operating Activity -'] = new_df['Cash from Operating Activity -'].str.replace(',', '')
        new_df['Cash from Operating Activity -'] = pd.to_numeric(new_df['Cash from Operating Activity -'])
        new_df['Cash from Investing Activity -'] = new_df['Cash from Investing Activity -'].str.replace(',', '')
        new_df['Cash from Investing Activity -'] = pd.to_numeric(new_df['Cash from Investing Activity -'])
        new_df['FCF'] = new_df['Cash from Operating Activity -'] + new_df['Cash from Investing Activity -']
        new_df['FCF'] = new_df['FCF'].fillna(0)

        year_count = 0
        for row in new_df.itertuples():
            if row.FCF < 0:
                year_count = year_count + 1
        if year_count > 0:
            a = a + 1
        return a
    
    ## Interest Coverage Ratio
    def Intrst(self, df, a):
        icr = df.iloc[0, 10]
        if icr == '':
            icr = 0
        else:
            icr = float(icr)
        if np.isnan(icr) == True:
            icr = np.nan_to_num(icr)

        if icr < 24 or icr > 100:
            a = a + 1
        return a
    
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

        year_count = 0
        for row in new_df.itertuples():
            if row.CFO_Zscore < -1 or row.CFO_Zscore > 1:
                year_count = year_count + 1
        n = int(0.75 * len(new_df))
        if year_count > n:
            a = a + 1

        return a
    
    ## Depreciation Rates
    def Depr(self, df, a):
        new_df = df[['Depreciation']]
        new_df['Depreciation'] = new_df['Depreciation'].str.replace(',', '')
        new_df['Depreciation'] = pd.to_numeric(new_df['Depreciation'])
        new_df['Difference in Depreciation'] = new_df['Depreciation'] - new_df['Depreciation'].shift(1)
        new_df.fillna(0, inplace = True)
        new_df['Depreciation Growth Rate'] = (new_df['Difference in Depreciation'] / new_df['Depreciation'].shift(1)) * 100
        new_df.fillna(0, inplace = True)

        mean = new_df['Depreciation Growth Rate'].mean()
        std = new_df['Depreciation Growth Rate'].std()
        new_df['Depreciation_Zscore'] = (new_df['Depreciation Growth Rate'] - mean) / std
        new_df['Depreciation_Zscore'] = new_df['Depreciation_Zscore'].fillna(0)

        year_count = 0
        for row in new_df.itertuples():
            if row.Depreciation_Zscore < -2 or row.Depreciation_Zscore > 2:
                year_count = year_count + 1
        if year_count > 0:
            a = a + 1

        return a
    
    ## Change in Reserves
    def Rsrv(self, df, a):
        new_df = df[['Reserves']]
        new_df['Reserves'] = new_df['Reserves'].str.replace(',', '')
        new_df['Reserves'] = pd.to_numeric(new_df['Reserves'])
        new_df['Difference'] = new_df['Reserves'] - new_df['Reserves'].shift(1)
        new_df.fillna(0, inplace = True)

        mean = new_df['Difference'].mean()
        std = new_df['Difference'].std()
        new_df['Reserves_Zscore'] = (new_df['Difference'] - mean) / std

        year_count = 0
        for row in new_df.itertuples():
            if row.Reserves_Zscore < -2 or row.Reserves_Zscore > 2:
                year_count = year_count + 1
        if year_count > 0:
            a = a + 1

        return a
    
    ## Yields on Cash and Cash Equivalents
    def Cash(self, df, a):
        new_df = df[['Interest']]
        new_df['Interest'] = new_df['Interest'].str.replace(',', '')
        new_df['Interest'] = pd.to_numeric(new_df['Interest'])
        new_df['Difference'] = new_df['Interest'] - new_df['Interest'].shift(1)
        new_df.fillna(0, inplace = True)

        mean = new_df['Difference'].mean()
        std = new_df['Difference'].std()
        new_df['Interest_Zscore'] = (new_df['Difference'] - mean) / std

        year_count = 0
        for row in new_df.itertuples():
            if row.Interest_Zscore < -1 or row.Interest_Zscore > 1:
                year_count = year_count + 1
        n = int(0.75 * len(new_df))
        if year_count > n:
            a = a + 1

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
            a = a + 1

        return a
    
    ## CWIP to Gross Block
    def CWIP(self, df, a):
        new_df = df[['CWIP', 'Gross Block']]
        new_df['CWIP'] = new_df['CWIP'].str.replace(',', '')
        new_df['CWIP'] = pd.to_numeric(new_df['CWIP'])
        new_df['Gross Block'] = new_df['Gross Block'].str.replace(',', '')
        new_df['Gross Block'] = pd.to_numeric(new_df['Gross Block'])
        new_df['Ratios'] = new_df['CWIP'] / new_df['Gross Block']
        new_df['Ratios'] = new_df['Ratios'].fillna(0)

        year_count = 0
        for row in new_df.itertuples():
            if row.Ratios < 0.05:
                year_count = year_count + 1
            elif row.Ratios > 0.75:
                year_count = year_count + 1
        n = int(0.75 * len(new_df))
        if year_count > n:
            a = a + 1

        return a
    
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

        error = 0
        error = self.ROA(transpose_df, error)
        error = self.Invt(top_ratios, error)
        error = self.CCC(top_ratios, market_leader, error)
        error = self.EPS(transpose_df, error)
        error = self.OCF(top_ratios, transpose_df, error)

        critical = 0
        critical = self.Rev_CAGR(transpose_df, critical)
        critical = self.Debt_Eq(top_ratios, critical)
        critical = self.ROCE(transpose_df, critical)
        critical = self.FCF(transpose_df, critical)
        critical = self.Intrst(top_ratios, critical)
        critical = self.CFO(transpose_df, critical)
        critical = self.Depr(transpose_df, critical)
        critical = self.Rsrv(transpose_df, critical)
        critical = self.Cash(transpose_df, critical)
        critical = self.Cont(top_ratios, critical)
        critical = self.CWIP(transpose_df, critical)
        
        str = ""
        if error > 2 and critical > 4:
            str = "Red Flag! Company not safe to Invest!"
        elif error <= 2 and critical > 4:
            str = "Orange Flag! Company needs further Investigation before Investing!"
        else:
            str = "Green Flag! Company safe to Invest!"
        
        return str

