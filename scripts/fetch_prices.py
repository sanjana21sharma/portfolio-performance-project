import yfinance as yf
import pandas as pd
import os

# --------- Configuration ---------
tickers = ['AAPL', 'MSFT', 'TSLA', 'GOOGL', 'AMZN']
start_date = '2015-04-01'
end_date = '2025-04-01'

# --------- Fetch Historical Prices ---------
print(f"Downloading data from {start_date} to {end_date} for: {', '.join(tickers)}")
raw_data = yf.download(tickers, start=start_date, end=end_date)

# --------- Extract Adjusted Close Prices ---------
data = raw_data['Close'].dropna()

# --------- Add Financial Year Column ---------
data = data.reset_index()

def get_fin_year(date):
    year = date.year
    return f"FY{year - 1}-{str(year)[-2:]}" if date.month < 4 else f"FY{year}-{str(year + 1)[-2:]}"
    
data['Financial Year'] = data['Date'].apply(get_fin_year)

# --------- Save to Excel ---------
output_path = r"C:\Users\anjxl\Desktop\Portfolio-Performance-Project\data\portfolio_data.xlsx"
data.to_excel(output_path, index=False)

print(f"✅ Data saved to: {output_path}")
