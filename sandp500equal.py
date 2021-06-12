import numpy as np
import pandas as pd
import requests
import xlsxwriter as xl
import math
from secrets import token


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

#Excel variables#
background_color = '#0a0a23'
font_color = '#ffffff'
#####################################

symbols = pd.read_csv('sp_500_stocks.csv')
symbol_groups = list(chunks(symbols['Ticker'], 100))
symbol_strings =[]
columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
portfolio_size = input('Enter the number value of your portfolio: ')
val = float(portfolio_size)


for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

final_dataframe = pd.DataFrame(columns = columns)

for symbol_string in symbol_strings:
    batch_api_url = f'https://cloud.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={token}'
    data = requests.get(batch_api_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe.append(
            pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    data[symbol]['quote']['marketCap'],
                    'N/A'
                ],
                index=columns),
                ignore_index=True
            )
position_size = val/len(final_dataframe.index)
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

string_format = writer.book.add_format(
    {
        'font_color' : font_color,
        'bg_color' : background_color,
        'border' : 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color' : font_color,
        'bg_color' : background_color,
        'border' : 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format' : '0',
        'font_color' : font_color,
        'bg_color' : background_color,
        'border' : 1
    }
)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]

}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Reccomended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])
writer.save()