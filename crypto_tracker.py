import requests
import pandas as pd
import time
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import PatternFill, Font
from generate_report import ReportGenerator

class CryptoTracker:
    def __init__(self):
        self.base_url = "https://api.coingecko.com/api/v3"
        self.excel_file = "crypto_live_data.xlsx"
        # Create report generator instance with the same excel file
        self.report_generator = ReportGenerator(self.excel_file)
        self.last_report_time = None
        # Set report interval to 5 minutes for testing (change back to 1 hour later)
        self.report_interval = timedelta(minutes=5)  

    def fetch_top_50_data(self):
        """Get cryptocurrency data from CoinGecko API"""
        try:
            # API endpoint for market data
            url = f"{self.base_url}/coins/markets"
            
            # Set query parameters
            params = {
                'vs_currency': 'usd',  # US Dollar as base currency
                'order': 'market_cap_desc',  # Sort by market cap
                'per_page': 50,  # Number of results
                'page': 1,
                'sparkline': False  # Don't need sparkline data
            }
            
            # Make API request with retry on failure
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    response = requests.get(url, params=params, timeout=10)
                    response.raise_for_status()
                    return response.json()
                except requests.RequestException as e:
                    if attempt == max_retries - 1:
                        print(f"Failed to fetch data after {max_retries} attempts: {e}")
                        return None
                    print(f"Attempt {attempt + 1} failed, retrying...")
                    time.sleep(2)
                    
        except Exception as e:
            print(f"Unexpected error: {e}")
            return None

    def process_data(self, data):
        """Clean and structure the raw API data"""
        if not data:
            return None
            
        try:
            # Convert to DataFrame and select needed columns
            df = pd.DataFrame(data)
            
            # Select required columns first
            cols_mapping = {
                'name': 'Name',
                'symbol': 'Symbol',
                'current_price': 'Price (USD)',
                'market_cap': 'Market Cap (USD)',
                'total_volume': '24h Volume (USD)',
                'price_change_percentage_24h': '24h Change (%)'
            }
            
            # Create a new DataFrame with only needed columns
            df = df[cols_mapping.keys()].copy()  # Create explicit copy
            
            # Fill NaN values without using inplace
            df['price_change_percentage_24h'] = df['price_change_percentage_24h'].fillna(0)
            df['total_volume'] = df['total_volume'].fillna(0)
            
            # Rename columns
            df = df.rename(columns=cols_mapping)
            
            # Convert symbols to uppercase
            df['Symbol'] = df['Symbol'].str.upper()
            
            return df
            
        except Exception as e:
            print(f"Error processing data: {e}")
            return None

    def analyze_data(self, df):
        """Perform analysis on the cryptocurrency data"""
        analysis = {
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'top_5_by_market_cap': df.head(5)['Name'].tolist(),
            'average_price': df['Price (USD)'].mean(),
            'highest_24h_change': df.nlargest(1, '24h Change (%)')['Name'].iloc[0],
            'lowest_24h_change': df.nsmallest(1, '24h Change (%)')['Name'].iloc[0],
            'highest_24h_change_value': df['24h Change (%)'].max(),
            'lowest_24h_change_value': df['24h Change (%)'].min()
        }
        return analysis

    def update_excel(self, df, analysis):
        """Update Excel file with latest data and analysis"""
        try:
            with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Live Data', index=False)
                
                # Write analysis
                analysis_df = pd.DataFrame({
                    'Metric': [
                        'Last Updated',
                        'Top 5 by Market Cap',
                        'Average Price (USD)',
                        'Highest 24h Change',
                        'Lowest 24h Change'
                    ],
                    'Value': [
                        analysis['timestamp'],
                        ', '.join(analysis['top_5_by_market_cap']),
                        f"${analysis['average_price']:.2f}",
                        f"{analysis['highest_24h_change']} ({analysis['highest_24h_change_value']:.2f}%)",
                        f"{analysis['lowest_24h_change']} ({analysis['lowest_24h_change_value']:.2f}%)"
                    ]
                })
                analysis_df.to_excel(writer, sheet_name='Analysis', index=False)
                
            print(f"Excel file updated successfully: {self.excel_file}")
            return True
        except Exception as e:
            print(f"Error updating Excel file: {e}")
            return False

    def run(self, update_interval=300):
        """Start the crypto tracker with periodic updates"""
        print("\n=== Starting Cryptocurrency Tracker ===")
        print(f"Data update interval: {update_interval} seconds")
        print(f"Report generation interval: {self.report_interval.total_seconds()/60:.1f} minutes")
        
        # Ensure directories exist
        self.report_generator.create_directories()
        
        # Generate initial report
        print("\nGenerating initial report...")
        success = self.report_generator.generate_report()
        if success:
            self.last_report_time = datetime.now()
            print("Initial report generated successfully")
        else:
            print("Failed to generate initial report")

        errors_count = 0
        max_errors = 5

        while True:
            try:
                current_time = datetime.now()
                print(f"\nFetching data at {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
                
                data = self.fetch_top_50_data()
                if not data:
                    errors_count += 1
                    if errors_count >= max_errors:
                        print("Too many consecutive errors. Stopping tracker.")
                        break
                    continue
                
                df = self.process_data(data)
                if df is not None:
                    analysis = self.analyze_data(df)
                    if self.update_excel(df, analysis):
                        print("Data updated successfully")
                        errors_count = 0

                        # Check if we should generate a new report
                        if not self.last_report_time or \
                           current_time - self.last_report_time >= self.report_interval:
                            print("\nGenerating new report...")
                            if self.report_generator.generate_report():
                                self.last_report_time = current_time
                                print("Report generated successfully")
                            else:
                                print("Failed to generate report")
                        else:
                            minutes_until_next = (self.report_interval - 
                                                (current_time - self.last_report_time)).total_seconds() / 60
                            print(f"Next report in {minutes_until_next:.1f} minutes")
                
                print(f"Next data update in {update_interval} seconds...")
                time.sleep(update_interval)
                
            except KeyboardInterrupt:
                print("\nStopping cryptocurrency tracker...")
                break
            except Exception as e:
                print(f"\nUnexpected error in main loop: {e}")
                errors_count += 1

if __name__ == "__main__":
    tracker = CryptoTracker()
    tracker.run() 