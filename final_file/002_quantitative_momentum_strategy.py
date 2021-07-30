import numpy as np
import pandas as pd
import math
import requests
import xlsxwriter
from scipy import stats
from scipy.stats import percentileofscore as score

#IMPORT OUR LIST OF STOCKS
stocks = pd.read_csv("C:\\V'S LIFE\\4_Project\\Algro-Trading\\actual_output\\Project 1\\sp_500_stocks.csv",'r')
#ACQUIRING AN API TOKEN
from secrets import IEX_CLOUD_API_TOKEN
#MAKING API CALL
symbol = 'AAPL'
api_url = f'https://sandbox.iexapis.com/stable/stock/{symbol}/quote/?token={IEX_CLOUD_API_TOKEN}'
data = requests.get(api_url).json()
#PARSING API
#data['year1ChangePercent']

#Executing A Batch API Call & Building Our DataFrame
# Function sourced from
# https://stackoverflow.com/questions/312443/how-do-you-split-a-list-into-evenly-sized-chunks
def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

my_columns = ['Ticker', 'Price', 'One-Year Price Return', 'Number of Shares to Buy']

final_dataframe = pd.DataFrame (columns=my_columns)
for symbol_string in symbol_strings:
    batch_api_call_url=f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=stats,quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data=requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe=final_dataframe.append(
        pd.Series([symbol,
                   data[symbol]['quote']['latestPrice'],
                   data[symbol]['stats']['year1ChangePercent'],
                   'N/A'], index = my_columns),
        ignore_index = True)

#REMOVING LOW MOMENTUM STOCKS
final_dataframe.sort_values('One-Year Price Return',ascending=False,inplace=True)
final_dataframe=final_dataframe[:50]
final_dataframe.reset_index(inplace=True)

#CALCULATING THE NUMBER OF SHARES TO BUY
def portfolio_input():
    global portfolio_size
    portfolio_size = input('Enter the size of your portfolio')

    try:
        float(portfolio_size)
    except:
        print('That is not a number')
        print('Please try again:')
        portfolio_size = input('Enter the size of your portfolio')
position_size=float(portfolio_size)/len(final_dataframe.index)
for i in range(0,len(final_dataframe)):
    final_dataframe.loc[i,'Number of Shares to Buy']=math.floor(position_size/final_dataframe.loc[i,'Price'])

#BUILDING A BETTER MOMENTUM STRATEGY
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
hqm_dataframe=pd.DataFrame(columns=hqm_columns)

for symbol_string in symbol_strings:
    batch_api_call_url=f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=stats,quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data=requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        hqm_dataframe = hqm_dataframe.append(
                                        pd.Series([symbol,
                                                   data[symbol]['quote']['latestPrice'],
                                                   'N/A',
                                                   data[symbol]['stats']['year1ChangePercent'],
                                                   'N/A',
                                                   data[symbol]['stats']['month6ChangePercent'],
                                                   'N/A',
                                                   data[symbol]['stats']['month3ChangePercent'],
                                                   'N/A',
                                                   data[symbol]['stats']['month1ChangePercent'],
                                                   'N/A',
                                                   'N/A'
                                                   ],
                                                  index = hqm_columns),
                                        ignore_index = True)

#CALCULATE MOMENTUM PERCENTILE
time_periods = [
                'One-Year',
                'Six-Month',
                'Three-Month',
                'One-Month'
                ]

for row in hqm_dataframe.index:
    for time_period in time_periods:
        hqm_dataframe.loc[row, f'{time_period} Return Percentile'] = stats.percentileofscore(hqm_dataframe[f'{time_period} Price Return'], hqm_dataframe.loc[row, f'{time_period} Price Return'])/100

# Print each percentile score to make sure it was calculated properly
for time_period in time_periods:
    print(hqm_dataframe[f'{time_period} Return Percentile'])

#CALCULATE HQM SCORE
from statistics import mean

for row in hqm_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(hqm_dataframe.loc[row, f'{time_period} Return Percentile'])
    hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)

#SELECTING SHARES TO BUY
hqm_dataframe.sort_values(by = 'HQM Score', ascending = False)
hqm_dataframe = hqm_dataframe[:51]

#CALCULATE THE NUMBER OF SHARES TO BUY
portfolio_input()
position_size = float(portfolio_size) / len(hqm_dataframe.index)
for i in range(0, len(hqm_dataframe['Ticker'])-1):
    hqm_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / hqm_dataframe['Price'][i])

#FORMATING EXCEL OUTPUT
writer = pd.ExcelWriter('momentum_strategy.xlsx', engine='xlsxwriter')
hqm_dataframe.to_excel(writer, sheet_name='Momentum Strategy', index = False)
background_color = '#0a0a23'
font_color = '#ffffff'

string_template = writer.book.add_format(
        {
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

dollar_template = writer.book.add_format(
        {
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_template = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

percent_template = writer.book.add_format(
        {
            'num_format':'0.0%',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

column_formats = {
                    'A': ['Ticker', string_template],
                    'B': ['Price', dollar_template],
                    'C': ['Number of Shares to Buy', integer_template],
                    'D': ['One-Year Price Return', percent_template],
                    'E': ['One-Year Return Percentile', percent_template],
                    'F': ['Six-Month Price Return', percent_template],
                    'G': ['Six-Month Return Percentile', percent_template],
                    'H': ['Three-Month Price Return', percent_template],
                    'I': ['Three-Month Return Percentile', percent_template],
                    'J': ['One-Month Price Return', percent_template],
                    'K': ['One-Month Return Percentile', percent_template],
                    'L': ['HQM Score', integer_template]
                    }

for column in column_formats.keys():
    writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Momentum Strategy'].write(f'{column}1', column_formats[column][0], string_template)
writer.save()
