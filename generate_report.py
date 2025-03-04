import pandas as pd
import os
from datetime import datetime, timedelta
import logging
from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Set up logging configuration
def setup_logging():
    """Configure logging with both file and console handlers"""
    # Create logs directory if it doesn't exist
    Path('logs').mkdir(exist_ok=True)
    
    # Create a formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Set up file handler
    file_handler = logging.FileHandler(
        os.path.join('logs', f'crypto_analysis_{datetime.now().strftime("%Y%m%d")}.log')
    )
    file_handler.setFormatter(formatter)
    
    # Set up console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    # Configure logger
    logger = logging.getLogger('CryptoAnalysis')
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Initialize logger
logger = setup_logging()

class ReportGenerator:
    def __init__(self, excel_file="crypto_live_data.xlsx"):
        self.excel_file = excel_file
        self.report_dir = "reports"
        self.archive_dir = os.path.join(self.report_dir, "archive")
        self.latest_report = os.path.join(self.report_dir, "latest_report.pdf")
        self.logger = logger
        
    def create_directories(self):
        """Ensure all necessary directories exist"""
        for directory in [self.report_dir, self.archive_dir]:
            Path(directory).mkdir(exist_ok=True)
            self.logger.info(f"Directory ready: {directory}")

    def load_data(self):
        """Load and validate data from Excel file"""
        try:
            if not os.path.exists(self.excel_file):
                self.logger.error(f"Excel file {self.excel_file} not found!")
                return None, None

            df = pd.read_excel(self.excel_file, sheet_name='Live Data')
            analysis_df = pd.read_excel(self.excel_file, sheet_name='Analysis')

            # Log successful data load
            self.logger.info(f"Successfully loaded data from {self.excel_file}")
            return df, analysis_df

        except Exception as e:
            self.logger.error(f"Error loading data: {str(e)}")
            return None, None

    def format_currency(self, value):
        """Format large numbers with appropriate suffixes"""
        if value >= 1e9:
            return f"${value/1e9:.2f}B"
        elif value >= 1e6:
            return f"${value/1e6:.2f}M"
        else:
            return f"${value:,.2f}"

    def create_pdf_report(self, df, analysis_df, output_path):
        """Create a PDF report using ReportLab"""
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        elements = []

        # Title
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            spaceAfter=30
        )
        elements.append(Paragraph("Cryptocurrency Market Analysis", title_style))
        elements.append(Spacer(1, 20))

        # Timestamp
        elements.append(Paragraph(
            f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            styles['Normal']
        ))
        elements.append(Spacer(1, 20))

        # Market Overview
        elements.append(Paragraph("Market Overview", styles['Heading2']))
        elements.append(Spacer(1, 10))

        # Create market overview table
        market_data = [
            ['Total Market Cap', self.format_currency(df['Market Cap (USD)'].sum())],
            ['Average Price', self.format_currency(df['Price (USD)'].mean())],
            ['24h Trading Volume', self.format_currency(df['24h Volume (USD)'].sum())]
        ]
        
        market_table = Table(market_data, colWidths=[200, 200])
        market_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(market_table)
        elements.append(Spacer(1, 20))

        # Top 5 Cryptocurrencies
        elements.append(Paragraph("Top 5 Cryptocurrencies by Market Cap", styles['Heading2']))
        elements.append(Spacer(1, 10))

        # Create top 5 table
        top_5_data = [['Name', 'Price (USD)', 'Market Cap (USD)', '24h Change (%)']]
        for _, row in df.head(5).iterrows():
            top_5_data.append([
                row['Name'],
                self.format_currency(row['Price (USD)']),
                self.format_currency(row['Market Cap (USD)']),
                f"{row['24h Change (%)']:.2f}%"
            ])

        top_5_table = Table(top_5_data, colWidths=[120, 100, 120, 100])
        top_5_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(top_5_table)

        # Build PDF
        doc.build(elements)

    def generate_report(self):
        """Generate comprehensive crypto market analysis report"""
        try:
            self.logger.info("Starting report generation...")
            self.create_directories()
            
            df, analysis_df = self.load_data()
            if df is None or analysis_df is None:
                return False

            # Generate timestamp for report
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Create archive report path
            archive_report = os.path.join(
                self.archive_dir, 
                f"crypto_analysis_{timestamp}.pdf"
            )

            # Generate PDF reports
            self.logger.info(f"Creating report: {self.latest_report}")
            self.create_pdf_report(df, analysis_df, self.latest_report)
            
            self.logger.info(f"Creating archived report: {archive_report}")
            self.create_pdf_report(df, analysis_df, archive_report)

            self.logger.info("Reports generated successfully")
            
            # Clean up old reports
            self._cleanup_old_reports()
            
            return True

        except Exception as e:
            self.logger.error(f"Failed to generate report: {str(e)}")
            return False

    def _cleanup_old_reports(self):
        """Remove reports older than 7 days"""
        try:
            current_time = datetime.now()
            for report in Path(self.archive_dir).glob("crypto_analysis_*.pdf"):
                timestamp_str = report.stem.split('_')[2:]
                report_time = datetime.strptime('_'.join(timestamp_str), "%Y%m%d_%H%M%S")
                
                if current_time - report_time > timedelta(days=7):
                    report.unlink()
                    self.logger.info(f"Removed old report: {report}")
        except Exception as e:
            self.logger.warning(f"Error cleaning up old reports: {e}")

def main():
    try:
        generator = ReportGenerator()
        success = generator.generate_report()
        if success:
            print("Report generated successfully!")
        else:
            print("Failed to generate report. Check the logs for details.")
    except KeyboardInterrupt:
        print("\nReport generation interrupted by user.")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")

if __name__ == "__main__":
    main() 