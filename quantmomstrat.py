import numpy as np
import pandas as pd
import requests
import math
from scipy import stats
import xlsxwriter
from secrets import token
from statistics import mean

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

hqm_columns = [
    'Ticker',
    'Price',
    'Number of Shares to Buy',
    'One-Year Price Return',
    'One-Year Return Percentile',
    'Six-Month Price Return',
    'Six-Month Return Percentile',
    'Three-Month Price Return',
    'Three-Month Return Percentile',
    'One-Month Price Return',
    'One-Month Return Percentile',
    'HQM Score'
]

final_dataframe = pd.DataFrame(columns = hqm_columns)
convert_none = lambda x : 0 if x is None else x

for symbol_string in symbol_strings:
    batch_api_url = f'https://cloud.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=stats,price,quote,news,chart&token={token}'
    data = requests.get(batch_api_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    'N/A',
                    convert_none(data[symbol]['stats']['year1ChangePercent']),
                    'N/A',
                    convert_none(data[symbol]['stats']['month6ChangePercent']),
                    'N/A',
                    convert_none(data[symbol]['stats']['month3ChangePercent']),
                    'N/A',
                    convert_none(data[symbol]['stats']['month1ChangePercent']),
                    'N/A',
                    'N/A'  
                    ],
                    index=hqm_columns
                ), ignore_index=True
            )
time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']

for row in final_dataframe.index:
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        final_dataframe.loc[row, percentile_col] = stats.percentileofscore(final_dataframe[change_col], final_dataframe.loc[row, change_col])/100

for row in final_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(final_dataframe.loc[row, f'{time_period} Return Percentile'])
    final_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)

final_dataframe.sort_values('HQM Score', ascending=False, inplace=True) 
final_dataframe = final_dataframe[:50]
final_dataframe.reset_index(drop=True, inplace=True)

position_size = float(portfolio_size)/len(final_dataframe.index)
for i in range(0, len(final_dataframe['Ticker'])-1):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe['Price'][i])

excel_dump(final_dataframe)