"""
# Excel Sales Analyzer

A Python tool for generating sample sales data, analyzing it, and creating professional Excel reports with charts and visualizations.

## Features
- Generates realistic sales data
- Creates formatted Excel files
- Performs sales analysis by product, region, and time
- Builds multi-sheet summary reports with charts
- Easy to integrate into existing data pipelines

## Requirements
- pandas
- numpy
- matplotlib
- xlsxwriter
- openpyxl

## Author
Ram Dudeja
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import random
import logging
import argparse
import os
import sys

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger('excel_sales_analyzer')


class SalesDataGenerator:
    """Class to generate sample sales data for demonstration"""
    
    def __init__(self):
        # Product categories and regions
        self.products = ["Laptop", "Desktop", "Monitor", "Keyboard", "Mouse", "Headphones", "Printer"]
        self.regions = ["North", "South", "East", "West", "Central"]
        self.sales_channels = ["Online", "Retail", "Distributor"]
        self.price_ranges = {
            "Laptop": (800, 2000),
            "Desktop": (600, 1800),
            "Monitor": (150, 500),
            "Keyboard": (20, 150),
            "Mouse": (10, 80),
            "Headphones": (30, 300),
            "Printer": (100, 400)
        }
    
    def generate_data(self, num_records=100, days_back=90):
        """
        Generate random sales data
        
        Args:
            num_records (int): Number of sales records to generate
            days_back (int): Number of days in the past to generate data for
            
        Returns:
            pandas.DataFrame: DataFrame containing the generated sales data
        """
        logger.info(f"Generating {num_records} sample sales records")
        
        # Create date range
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days_back)
        dates = [start_date + timedelta(days=x) for x in range((end_date - start_date).days)]
        
        # Generate random data
        sales_data = []
        for _ in range(num_records):
            product = random.choice(self.products)
            unit_price = round(random.uniform(*self.price_ranges[product]), 2)
            quantity = random.randint(1, 10)
            date = random.choice(dates)
            region = random.choice(self.regions)
            channel = random.choice(self.sales_channels)
            
            sales_data.append({
                "Date": date.strftime("%Y-%m-%d"),
                "Product": product,
                "Region": region,
                "Channel": channel,
                "Units": quantity,
                "Unit_Price": unit_price,
                "Total_Sale": round(quantity * unit_price, 2)
            })
        
        return pd.DataFrame(sales_data)


class ExcelWriter:
    """Class to write and format Excel files"""
    
    def write_sales_data(self, df, filename="sales_data.xlsx", output_dir="output"):
        """
        Write the sales DataFrame to an Excel file with proper formatting
        
        Args:
            df (pandas.DataFrame): Sales data to write
            filename (str): Name of the Excel file
            output_dir (str): Directory to save the file
            
        Returns:
            str: Path to the created Excel file
        """
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        filepath = os.path.join(output_dir, filename)
        
        logger.info(f"Writing data to {filepath}")
        
        try:
            # Create Excel writer
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                # Write data to sheet
                df.to_excel(writer, sheet_name='Sales_Data', index=False)
                
                # Get the workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets['Sales_Data']
                
                # Add formats
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#D7E4BC',
                    'border': 1
                })
                
                money_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
                date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
                border_format = workbook.add_format({'border': 1})
                
                # Apply formats
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Format the columns
                worksheet.set_column('A:A', 12, date_format)  # Date column
                worksheet.set_column('B:D', 15, border_format)  # Product, Region, Channel
                worksheet.set_column('E:E', 8, border_format)  # Units
                worksheet.set_column('F:G', 12, money_format)  # Price and Total columns
                
                # Add auto-filter
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
            logger.info(f"Data successfully written to {filepath}")
            return filepath
        
        except Exception as e:
            logger.error(f"Error writing Excel file: {str(e)}")
            raise


class SalesAnalyzer:
    """Class to analyze sales data"""
    
    def __init__(self, df=None):
        """
        Initialize the analyzer with optional data
        
        Args:
            df (pandas.DataFrame, optional): Sales data to analyze
        """
        self.df = df
    
    def load_from_excel(self, filename):
        """
        Load sales data from Excel file
        
        Args:
            filename (str): Path to the Excel file
            
        Returns:
            SalesAnalyzer: Self for method chaining
        """
        try:
            logger.info(f"Loading data from {filename}")
            self.df = pd.read_excel(filename)
            return self
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            raise
    
    def analyze(self):
        """
        Perform analysis on the sales data
        
        Returns:
            dict: Dictionary containing analysis results
        """
        if self.df is None:
            logger.error("No data available for analysis")
            raise ValueError("No data available for analysis")
        
        logger.info("Performing sales data analysis")
        
        # Ensure date column is datetime
        self.df['Date'] = pd.to_datetime(self.df['Date'])
        
        # Basic statistics and analysis
        total_sales = self.df['Total_Sale'].sum()
        avg_sale_per_transaction = self.df['Total_Sale'].mean()
        max_sale = self.df['Total_Sale'].max()
        
        # Sales by product
        product_sales = self.df.groupby('Product')['Total_Sale'].sum().sort_values(ascending=False)
        
        # Sales by region
        region_sales = self.df.groupby('Region')['Total_Sale'].sum().sort_values(ascending=False)
        
        # Sales by channel
        channel_sales = self.df.groupby('Channel')['Total_Sale'].sum().sort_values(ascending=False)
        
        # Sales trend over time
        self.df['Week'] = self.df['Date'].dt.isocalendar().week
        weekly_sales = self.df.groupby('Week')['Total_Sale'].sum()
        
        # Units sold by product
        units_by_product = self.df.groupby('Product')['Units'].sum().sort_values(ascending=False)
        
        # Average sale by region
        avg_sale_by_region = self.df.groupby('Region')['Total_Sale'].mean().sort_values(ascending=False)
        
        # Return all analyses
        return {
            "total_sales": total_sales,
            "avg_sale": avg_sale_per_transaction,
            "max_sale": max_sale,
            "product_sales": product_sales,
            "region_sales": region_sales,
            "channel_sales": channel_sales,
            "weekly_sales": weekly_sales,
            "units_by_product": units_by_product,
            "avg_sale_by_region": avg_sale_by_region,
            "raw_data": self.df
        }


class ReportGenerator:
    """Class to generate summary reports from analysis results"""
    
    def create_report(self, analysis_results, filename="sales_summary.xlsx", output_dir="output"):
        """
        Create a summary report in Excel with charts and tables
        
        Args:
            analysis_results (dict): Results from the SalesAnalyzer
            filename (str): Name of the Excel file
            output_dir (str): Directory to save the file
            
        Returns:
            str: Path to the created report file
        """
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        filepath = os.path.join(output_dir, filename)
        
        logger.info(f"Creating summary report at {filepath}")
        
        try:
            # Create Excel writer
            with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Create Summary Sheet
                self._create_summary_sheet(workbook, writer, analysis_results)
                
                # Create Product Sales Sheet
                self._create_product_sheet(workbook, writer, analysis_results)
                
                # Create Regional Sales Sheet
                self._create_region_sheet(workbook, writer, analysis_results)
                
                # Create Weekly Sales Trend Sheet
                self._create_trend_sheet(workbook, writer, analysis_results)
                
                # Create Channel Sheet
                self._create_channel_sheet(workbook, writer, analysis_results)
            
            logger.info(f"Summary report successfully created at {filepath}")
            return filepath
        
        except Exception as e:
            logger.error(f"Error creating summary report: {str(e)}")
            raise
    
    def _create_summary_sheet(self, workbook, writer, analysis_results):
        """Create the summary overview sheet"""
        summary_data = pd.DataFrame({
            'Metric': [
                'Total Sales', 
                'Average Sale per Transaction', 
                'Maximum Sale',
                'Top Product',
                'Top Region',
                'Top Channel'
            ],
            'Value': [
                f"${analysis_results['total_sales']:,.2f}",
                f"${analysis_results['avg_sale']:,.2f}",
                f"${analysis_results['max_sale']:,.2f}",
                f"{analysis_results['product_sales'].index[0]}",
                f"{analysis_results['region_sales'].index[0]}",
                f"{analysis_results['channel_sales'].index[0]}"
            ]
        })
        
        summary_data.to_excel(writer, sheet_name='Summary', index=False)
        summary_sheet = writer.sheets['Summary']
        
        # Format summary sheet
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#B8CCE4',
            'border': 1,
            'align': 'center'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'align': 'left'
        })
        
        for col_num, value in enumerate(summary_data.columns.values):
            summary_sheet.write(0, col_num, value, header_format)
        
        # Apply cell format to all cells
        for row_num in range(1, len(summary_data) + 1):
            for col_num in range(len(summary_data.columns)):
                summary_sheet.write(row_num, col_num, summary_data.iloc[row_num-1, col_num], cell_format)
        
        summary_sheet.set_column('A:A', 30)
        summary_sheet.set_column('B:B', 20)
    
    def _create_product_sheet(self, workbook, writer, analysis_results):
        """Create the product sales sheet with chart"""
        product_data = analysis_results['product_sales'].reset_index()
        product_data.columns = ['Product', 'Total Sales']
        product_data['Total Sales'] = product_data['Total Sales'].round(2)
        
        product_data.to_excel(writer, sheet_name='Product_Sales', index=False)
        product_sheet = writer.sheets['Product_Sales']
        
        # Format columns
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#E6B8B7',
            'border': 1,
            'align': 'center'
        })
        
        for col_num, value in enumerate(product_data.columns.values):
            product_sheet.write(0, col_num, value, header_format)
        
        # Add chart
        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Sales by Product',
            'categories': ['Product_Sales', 1, 0, len(product_data), 0],
            'values': ['Product_Sales', 1, 1, len(product_data), 1],
            'data_labels': {'value': True}
        })
        
        chart.set_title({'name': 'Sales by Product'})
        chart.set_x_axis({'name': 'Product'})
        chart.set_y_axis({'name': 'Sales ($)'})
        chart.set_style(11)  # Add a style
        chart.set_size({'width': 720, 'height': 400})
        product_sheet.insert_chart('D2', chart)
    
    def _create_region_sheet(self, workbook, writer, analysis_results):
        """Create the regional sales sheet with chart"""
        region_data = analysis_results['region_sales'].reset_index()
        region_data.columns = ['Region', 'Total Sales']
        region_data['Total Sales'] = region_data['Total Sales'].round(2)
        
        # Add percentage column
        total = region_data['Total Sales'].sum()
        region_data['Percentage'] = (region_data['Total Sales'] / total * 100).round(2)
        region_data['Percentage'] = region_data['Percentage'].apply(lambda x: f"{x}%")
        
        region_data.to_excel(writer, sheet_name='Regional_Sales', index=False)
        region_sheet = writer.sheets['Regional_Sales']
        
        # Format header
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#B7DEE8',
            'border': 1,
            'align': 'center'
        })
        
        for col_num, value in enumerate(region_data.columns.values):
            region_sheet.write(0, col_num, value, header_format)
        
        # Add pie chart
        chart = workbook.add_chart({'type': 'pie'})
        chart.add_series({
            'name': 'Sales by Region',
            'categories': ['Regional_Sales', 1, 0, len(region_data), 0],
            'values': ['Regional_Sales', 1, 1, len(region_data), 1],
            'data_labels': {'percentage': True}
        })
        
        chart.set_title({'name': 'Sales by Region'})
        chart.set_style(10)
        chart.set_size({'width': 600, 'height': 400})
        region_sheet.insert_chart('E2', chart)
    
    def _create_trend_sheet(self, workbook, writer, analysis_results):
        """Create the weekly trend sheet with chart"""
        weekly_data = analysis_results['weekly_sales'].reset_index()
        weekly_data.columns = ['Week', 'Total Sales']
        weekly_data['Total Sales'] = weekly_data['Total Sales'].round(2)
        
        weekly_data.to_excel(writer, sheet_name='Weekly_Trend', index=False)
        trend_sheet = writer.sheets['Weekly_Trend']
        
        # Format header
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#CCC0DA',
            'border': 1,
            'align': 'center'
        })
        
        for col_num, value in enumerate(weekly_data.columns.values):
            trend_sheet.write(0, col_num, value, header_format)
        
        # Add line chart
        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({
            'name': 'Weekly Sales Trend',
            'categories': ['Weekly_Trend', 1, 0, len(weekly_data), 0],
            'values': ['Weekly_Trend', 1, 1, len(weekly_data), 1],
            'marker': {'type': 'circle'},
            'line': {'width': 2.5}
        })
        
        chart.set_title({'name': 'Weekly Sales Trend'})
        chart.set_x_axis({'name': 'Week Number'})
        chart.set_y_axis({'name': 'Total Sales ($)'})
        chart.set_style(12)
        chart.set_size({'width': 720, 'height': 400})
        trend_sheet.insert_chart('D2', chart)
    
    def _create_channel_sheet(self, workbook, writer, analysis_results):
        """Create the sales channel sheet with chart"""
        channel_data = analysis_results['channel_sales'].reset_index()
        channel_data.columns = ['Channel', 'Total Sales']
        channel_data['Total Sales'] = channel_data['Total Sales'].round(2)
        
        channel_data.to_excel(writer, sheet_name='Channel_Sales', index=False)
        channel_sheet = writer.sheets['Channel_Sales']
        
        # Format header
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#D8E4BC',
            'border': 1,
            'align': 'center'
        })
        
        for col_num, value in enumerate(channel_data.columns.values):
            channel_sheet.write(0, col_num, value, header_format)
        
        # Add bar chart
        chart = workbook.add_chart({'type': 'bar'})
        chart.add_series({
            'name': 'Sales by Channel',
            'categories': ['Channel_Sales', 1, 0, len(channel_data), 0],
            'values': ['Channel_Sales', 1, 1, len(channel_data), 1],
            'data_labels': {'value': True}
        })
        
        chart.set_title({'name': 'Sales by Channel'})
        chart.set_x_axis({'name': 'Channel'})
        chart.set_y_axis({'name': 'Sales ($)'})
        chart.set_style(11)
        chart.set_size({'width': 600, 'height': 400})
        channel_sheet.insert_chart('D2', chart)


def parse_args():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description='Excel Sales Data Generator and Analyzer')
    parser.add_argument('--records', type=int, default=150, help='Number of sample records to generate')
    parser.add_argument('--days', type=int, default=90, help='Number of days in the past to generate data for')
    parser.add_argument('--output-dir', type=str, default='output', help='Directory to save output files')
    parser.add_argument('--data-file', type=str, default='sales_data.xlsx', help='Filename for raw data')
    parser.add_argument('--report-file', type=str, default='sales_summary.xlsx', help='Filename for summary report')
    parser.add_argument('--verbose', action='store_true', help='Enable verbose logging')
    
    return parser.parse_args()


def main():
    """Main execution function"""
    # Parse command line arguments
    args = parse_args()
    
    # Set logging level based on verbose flag
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    try:
        # Step 1: Generate sales data
        logger.info("Starting sales data generation and analysis process")
        data_generator = SalesDataGenerator()
        sales_df = data_generator.generate_data(num_records=args.records, days_back=args.days)
        
        # Step 2: Write to Excel
        excel_writer = ExcelWriter()
        sales_file = excel_writer.write_sales_data(
            sales_df, 
            filename=args.data_file, 
            output_dir=args.output_dir
        )
        
        # Step 3: Analyze sales data
        analyzer = SalesAnalyzer(sales_df)
        analysis_results = analyzer.analyze()
        
        # Step 4: Create summary report
        report_generator = ReportGenerator()
        report_file = report_generator.create_report(
            analysis_results, 
            filename=args.report_file, 
            output_dir=args.output_dir
        )
        
        # Print key insights
        logger.info("\nProcess completed successfully!")
        logger.info(f"Raw data file: {sales_file}")
        logger.info(f"Summary report: {report_file}")
        
        logger.info("\nKey Insights:")
        logger.info(f"Total Sales: ${analysis_results['total_sales']:,.2f}")
        logger.info(f"Best-selling product: {analysis_results['product_sales'].index[0]} " +
                   f"(${analysis_results['product_sales'].iloc[0]:,.2f})")
        logger.info(f"Top-performing region: {analysis_results['region_sales'].index[0]} " +
                   f"(${analysis_results['region_sales'].iloc[0]:,.2f})")
        logger.info(f"Best sales channel: {analysis_results['channel_sales'].index[0]} " +
                   f"(${analysis_results['channel_sales'].iloc[0]:,.2f})")
        
        return 0
    
    except Exception as e:
        logger.error(f"Error in main execution: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
