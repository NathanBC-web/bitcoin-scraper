from pycoingecko import CoinGeckoAPI
import pandas as pd
from datetime import datetime
import os
import logging

# Configure logging
logging.basicConfig(
    filename='bitcoin_scraper.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

try:
    logging.info("Starting Bitcoin price scraping process...")
    cg = CoinGeckoAPI()

    # Get Bitcoin price data for the last 90 days
    logging.info("Fetching Bitcoin price data from CoinGecko...")
    data = cg.get_coin_market_chart_by_id(id='bitcoin', vs_currency='usd', days=90)

    # Convert timestamps to readable dates and collect prices
    logging.info("Processing price data into DataFrame...")
    prices = [(datetime.fromtimestamp(p[0]/1000).strftime('%Y-%m-%d'), p[1]) for p in data['prices']]
    df = pd.DataFrame(prices, columns=['Date', 'Bitcoin Price'])

    # Group by date and calculate the average price per day
    logging.info("Calculating daily average prices...")
    daily_avg_df = df.groupby('Date', as_index=False).mean()

    # Sort by date descending (newest first)
    logging.info("Sorting data by date (newest first)...")
    daily_avg_df = daily_avg_df.sort_values(by='Date', ascending=False)

    # Add empty Expected Bitcoin Price column
    logging.info("Adding empty Expected Bitcoin Price column...")
    daily_avg_df['Expected Bitcoin Price'] = None

    # Define output path with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = f'C:\\Users\\ArhanPeek\\OneDrive - Universal IT B.V\\Documents\\Crypto\\bitcoin_prices_{timestamp}.xlsx'

    # Create folder if it doesn't exist
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Save to Excel with column width and frozen header
    logging.info(f"Saving data to Excel file: {output_path}")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        daily_avg_df.to_excel(writer, sheet_name='Bitcoin Prices', index=False)
        
        # Adjust column widths
        workbook = writer.book
        worksheet = writer.sheets['Bitcoin Prices']
        for column in ['Date', 'Bitcoin Price', 'Expected Bitcoin Price']:
            max_length = max(
                len(str(column)),
                daily_avg_df[column].astype(str).str.len().max() if column != 'Expected Bitcoin Price' else len(column)
            )
            worksheet.column_dimensions[column[0]].width = max_length + 2  # Add padding

        # Freeze the header row (row 1)
        worksheet.freeze_panes = 'A2'

    print(f"✅ Daily average price Excel file saved successfully at:\n{output_path}")
    logging.info("Bitcoin price scraping process completed successfully.")

except Exception as e:
    error_message = f"An error occurred: {str(e)}"
    print("❌ " + error_message)
    logging.error(error_message)

# Pause so the window stays open
input("\nPress Enter to close...")