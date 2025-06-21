import pandas as pd
import numpy as np
import os

# --------- Load Excel Data ---------
input_path = r"C:\Users\anjxl\Desktop\Portfolio-Performance-Project\data\portfolio_data.xlsx"
df = pd.read_excel(input_path)

# --------- Prepare Data ---------
df['Date'] = pd.to_datetime(df['Date'])
df.set_index('Date', inplace=True)

# Drop non-price columns
price_df = df.drop(columns=['Financial Year'])

# --------- Calculate Daily Returns ---------
returns_df = price_df.pct_change().dropna()

# --------- Define Metric Functions ---------
def annualized_return(daily_returns):
    return daily_returns.mean() * 252

def annualized_volatility(daily_returns):
    return daily_returns.std() * np.sqrt(252)

def sharpe_ratio(daily_returns, risk_free_rate=0.04):
    excess_return = daily_returns.mean() - (risk_free_rate / 252)
    return (excess_return / daily_returns.std()) * np.sqrt(252)

# --------- Calculate Metrics Per Stock ---------
metrics = pd.DataFrame(index=returns_df.columns)
metrics['Annual Return (%)'] = annualized_return(returns_df) * 100
metrics['Volatility (%)'] = annualized_volatility(returns_df) * 100
metrics['Sharpe Ratio'] = sharpe_ratio(returns_df)

# --------- Round and Save ---------
metrics = metrics.round(2)
output_path = r"C:\Users\anjxl\Desktop\Portfolio-Performance-Project\reports\portfolio_metrics.xlsx"
os.makedirs(os.path.dirname(output_path), exist_ok=True)
metrics.to_excel(output_path)

print("✅ Portfolio metrics calculated and saved to:")
print(output_path)
