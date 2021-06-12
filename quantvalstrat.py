import numpy as np
import pandas as pd
from scipy.stats.stats import PointbiserialrResult
import xlsxwriter
import requests
from scipy import stats
import math
from statistics import mean
from secrets import token


convert_none = lambda x : 0 if x is None else x

def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def get_symbols():
    symbols = []
    data = requests.get(f'https://cloud.iexapis.com/beta/ref-data/symbols?token={token}').json()
    for da in data:
        symbols.append(da['symbol'])
    return symbols

def portfolio_input():
    global portfolio_size
    portfolio_size = input('Enter the size of your portfolio: ')
    try:
        float(portfolio_size)
    except ValueError:
        print('Not a number')

def excel_dump(final_dataframe):
    writer = pd.ExcelWriter('momentum_strategy.xlsx', engine='xlsxwriter')
    final_dataframe.to_excel(writer, sheet_name='Momentum Strategy', index=False)
    background_color = '#000000'
    font_color = "#ffffff"

    string_template = writer.book.add_format(
        {
            'font_color': '#ffffff',
            'bg_color' : '#408ec6',
            'border' : 1
        }
    )

    dollar_template = writer.book.add_format(
        {
            'num_format' : '$0.00',
            'font_color' : '#ffffff',
            'bg_color' : '#408ec6',
            'border' : 1
        }
    )

    integer_template = writer.book.add_format(
        {
            'font_color' : '#ffffff',
            'bg_color' : '#408ec6',
            'border' : 1
        }
    )

    price_return_template = writer.book.add_format(
        {
            'num_format' : '0.0%',
            'font_color' : '#ffffff',
            'bg_color' : '#7a2048',
            'border' : 1
        }
    )

    return_percentile_template = writer.book.add_format(
        {
            'num_format' : '0.0%',
            'font_color' : '#ffffff',
            'bg_color' : '1e2761',
            'border' : 1
        }
    )

    column_formats = {
        'A': ['Ticker', string_template],
        'B': ['Price', dollar_template],
        'C': ['Number of Shares to Buy', integer_template],
        'D': ['One-Year Price Return', price_return_template],
        'E': ['One-Year Return Percentile', return_percentile_template],
        'F': ['Six-Month Price Return', price_return_template],
        'G': ['Six-Month Return Percentile', return_percentile_template],
        'H': ['Three-Month Price Return', price_return_template],
        'I': ['Three-Month Return Percentile', return_percentile_template],
        'J': ['One-Month Price Return', price_return_template],
        'K': ['One-Month Return Percentile', return_percentile_template],
        'L': ['HQM Score', integer_template]
    }

    for column in column_formats.keys():
        writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 20, column_formats[column][1])
        writer.sheets['Momentum Strategy'].write(f'{column}1', column_formats[column][0], string_template)
    
    writer.save()

portfolio_input()
#symbols = get_symbols()
symbols = pd.read_csv('sp_500_stocks.csv')
#symbol_groups = list(chunks(symbols, 100))
symbol_groups = list(chunks(symbols['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

rv_columns = [                                      
    'Ticker',
    'Price', 
    'Price-to-Earnings Ratio',
    'PE Percentile', 
    'Price-to-Book Ratio',
    'PB Percentile', 
    'Price-to-Sales Ratio',
    'PS Percentile', 
    'EV/EBITDA',
    'EV/EBITDA Percentile',
    'EV/GP',
    'EV/GP Percentile',
    'RV Score', 
    'Number of Shares to buy'
    ]

final_dataframe = pd.DataFrame(columns = rv_columns)

for symbol_string in symbol_strings:
    batch_api_url = f'https://cloud.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote,advanced-stats&token={token}'
    data = requests.get(batch_api_url).json()
    print(data)
    for symbol in symbol_string.split(','):
        price = data[symbol]['quote']['latestPrice']
        pe_ratio = data[symbol]['quote']['peRatio']
        pb_ratio = data[symbol]['advanced-stats']['priceToBook']
        ps_ratio = data[symbol]['advanced-stats']['priceToSales']
        enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']
        ebitda = data[symbol]['advanced-stats']['EBITDA']
        gross_profit = data[symbol]['advanced-stats']['grossProfit']
        
        try:
            ev_to_ebitda = enterprise_value/ebitda
        except TypeError:
            ev_to_ebitda = np.NaN

        try:
            ev_to_gross_profit = enterprise_value/gross_profit
        except TypeError:
            ev_to_gross_profit = np.NaN

        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    symbol,
                    price,
                    pe_ratio,
                    'N/A',
                    pb_ratio,
                    'N/A',
                    ps_ratio,
                    'N/A',
                    ev_to_ebitda,
                    'N/A',
                    ev_to_gross_profit,
                    'N/A',
                    'N/A',
                    'N/A'
                ], index = rv_columns
            ), ignore_index = True
        )

for column in ['Price-to-Earnings Ratio', 'Price-to-Book Ratio', 'Price-to-Sales Ratio', 'EV/EBITDA', 'EV/GP']:
    final_dataframe[column].fillna(final_dataframe[column].mean(), inplace = True)

convert_none = lambda x : 0 if x is None else x

metrics = {
    'Price-to-Earnings Ratio' : 'PE Percentile', 
    'Price-to-Book Ratio' : 'PB Percentile', 
    'Price-to-Sales Ratio' : 'PS Percentile', 
    'EV/EBITDA' : 'EV/EBITDA Percentile',
    'EV/GP' : 'EV/GP Percentile',
}

for metric in metrics.keys():
    for row in final_dataframe.index:
        final_dataframe.loc[row, metrics[metric]] = stats.percentileofscore(final_dataframe[metric], final_dataframe.loc[row, metrics[metric]])

for row in final_dataframe.index:
    value_percentiles = []
    for metric in metrics.keys():
        value_percentiles.append(final_dataframe.loc[row, metrics[metric]])
    final_dataframe.loc[row, 'RV Score'] = mean(value_percentiles)

final_dataframe.sort_values('Price-to-Earnings Ratio', ascending = False, inplace = True)
final_dataframe = final_dataframe[final_dataframe['Price-to-Earnings Ratio'] > 0 ]
final_dataframe = final_dataframe[:50]
final_dataframe.drop('index', axis = 1, inplace = True)

position_size = float(portfolio_size)/len(final_dataframe.index)
for i in range(0, len(final_dataframe['Ticker'])-1):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe['Price'][i])