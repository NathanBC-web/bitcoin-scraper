from pycoingecko import CoinGeckoAPI
import pandas as pd
from datetime import datetime
import os

try:
    cg = CoinGeckoAPI()

    # Get Bitcoin price data for the last 30 days (includes hourly data)
    data = cg.get_coin_market_chart_by_id(id='bitcoin', vs_currency='usd', days=90)

    # Convert timestamps to readable dates and collect prices
    prices = [(datetime.fromtimestamp(p[0]/1000).strftime('%Y-%m-%d'), p[1]) for p in data['prices']]
    df = pd.DataFrame(prices, columns=['Date', 'Bitcoin Price'])

    # Group by date and calculate the average price per day
    daily_avg_df = df.groupby('Date', as_index=False).mean()

    # Sort by date descending (newest first)
    daily_avg_df = daily_avg_df.sort_values(by='Date', ascending=False)

    # Define output path
    from datetime import datetime
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = f'C:\\Users\\ArhanPeek\\OneDrive - Universal IT B.V\\Documents\\Crypto\\bitcoin_prices_{timestamp}.xlsx'

    # Create folder if it doesn't exist
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Save to Excel (overwrite each time)
    daily_avg_df.to_excel(output_path, index=False)

    print(f"✅ Daily average price Excel file saved successfully at:\n{output_path}")

except Exception as e:
    print("❌ An error occurred:")
    print(e)

# Pause so the window stays open
input("\nPress Enter to close...")
