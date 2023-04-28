
#!/usr/bin/env python
# coding: utf-8

# In[1]:


from flask import Flask, jsonify, request
from flask_cors import CORS


# In[2]:


from StockScreener import Stock_Market
from TopStocks import Top_Stock
from Metrics import Stock_Metric


# In[3]:


import pandas as pd


# In[4]:


stock_symbol = pd.read_csv('Stock Symbol List.csv', index_col = 0)


# In[5]:


app = Flask(__name__)

cors = CORS(app, resources = {r"*" :{"origins" :"*"}})

# In[6]:


@app.route('/api/stock', methods = ['GET'])
def stock_api():
    # Get the data from the request
    data = request.get_json()["CompanyName"]
    print(data)

    # Check if company exists in stock_symbol DataFrame
    if data not in stock_symbol.index:
        return jsonify({'error': f'The company "{data}" is not a part of the current list.'})

    # Process the data using your class
    obj = Stock_Market()
    row = stock_symbol.loc[data].to_list()
    result_str = obj.screener(row)

    # Return a response
    return jsonify({'result': result_str})

@app.route('/api/top', methods = ['POST'])
def top_api():

    # Process the data using your class
    top = pd.DataFrame({'Company': pd.Series(dtype='str'),
                        'Error': pd.Series(dtype='int'),
                        'Critical': pd.Series(dtype='int')})

    for index in stock_symbol.index:
        obj = Top_Stock()
        row = stock_symbol.loc[index].to_list()
        str1, err, cri = obj.screener(row)
        if str1 == "Green Flag! Company safe to Invest!":
            top.loc[len(top.index)] = [index, err, cri]
    top = top.sort_values(by=['Critical', 'Error'], ascending=[True, True])
    top = top.head(10)
    lst = top.values.tolist()

    # Return a response
    return jsonify({'result': lst})

@app.route('/api/metric', methods = ['GET'])
def metric_api():
    # Get the data from the request
    data = request.get_json()["CompanyName"]
    print(data)

    # Check if company exists in stock_symbol DataFrame
    if data not in stock_symbol.index:
        return jsonify({'error': f'The company "{data}" is not a part of the current list.'})

    # Process the data using your class
    obj = Stock_Metric()
    row = stock_symbol.loc[data].to_list()
    result_str = obj.screener(row)

    # Return a response
    return jsonify({'result': result_str})
# In[7]:


if __name__ == '__main__':
    # app.run(debug = True)
    app.run("0.0.0.0", port = 5000, debug=True)
