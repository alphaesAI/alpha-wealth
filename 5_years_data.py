import yfinance as yf
import pandas as pd
from pymongo import MongoClient
from datetime import datetime

# Configuration
ticker = 'TCS.NS'  # TCS on NSE
start_date = (datetime.now() - pd.DateOffset(years=5)).strftime('%Y-%m-%d')
end_date = datetime.now().strftime('%Y-%m-%d')
mongo_uri = "mongodb://localhost:27017/"  # Replace with your MongoDB URI if needed
database_name = "stock_data"
collection_name = "tcs_stock_5_years"

# Fetch stock data
data = yf.download(ticker, start=start_date, end=end_date)

# Reset index to convert Date index to a column for MongoDB compatibility
data.reset_index(inplace=True)

# Store data in MongoDB
client = MongoClient(mongo_uri)
db = client[database_name]
collection = db[collection_name]

# Convert the DataFrame to a list of dictionaries and insert into MongoDB
data_dict = data.to_dict("records")
collection.insert_many(data_dict)
print(f"Inserted {len(data_dict)} records into MongoDB collection '{collection_name}'.")

# Save data to an Excel file
excel_file = "tcs_stock_5_years.xlsx"
data.to_excel(excel_file, index=False)
print(f"Data saved to Excel file: {excel_file}")

# Close MongoDB connection
client.close()
