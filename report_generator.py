#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Automated Report Generation Script
This script reads data from a CSV file, analyzes it, and generates a formatted PDF report.
"""

import csv
import os
import sys
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np

try:
    from fpdf import FPDF
except ImportError:
    print("FPDF library not found. Installing...")
    try:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "fpdf"])
        from fpdf import FPDF
        print("FPDF installed successfully.")
    except Exception as e:
        print(f"Failed to install FPDF: {e}")
        sys.exit(1)

def create_sample_data(filename):
    """Create sample CSV data if file doesn't exist"""
    if os.path.exists(filename):
        print(f"Data file '{filename}' already exists.")
        return
    
    sample_data = """Date,Product,Region,Sales,Expenses
2023-01-01,Product A,North,5000,3000
2023-01-01,Product B,North,4500,2800
2023-01-01,Product C,North,6000,3500
2023-01-01,Product A,South,5500,3200
2023-01-01,Product B,South,4800,2900
2023-01-01,Product C,South,6200,3800
2023-02-01,Product A,North,5200,3100
2023-02-01,Product B,North,4700,2850
2023-02-01,Product C,North,6100,3600
2023-02-01,Product A,South,5600,3300
2023-02-01,Product B,South,4900,2950
2023-02-01,Product C,South,6300,3900
2023-03-01,Product A,North,5300,3150
2023-03-01,Product B,North,4800,2900
2023-03-01,Product C,North,6200,3650
2023-03-01,Product A,South,5700,3350
2023-03-01,Product B,South,5000,3000
2023-03-01,Product C,South,6400,3950"""
    
    with open(filename, 'w') as f:
        f.write(sample_data)
    
    print(f"Sample data file '{filename}' created successfully.")

def read_and_analyze_data(filename):
    """Read data from CSV file and perform basic analysis"""
    data = []
    try:
        with open(filename, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                data.append(row)
    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
        return None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None
    
    # Convert numeric fields to appropriate types
    for row in data:
        try:
            row['Sales'] = float(row['Sales'])
            row['Expenses'] = float(row['Expenses'])
            row['Profit'] = row['Sales'] - row['Expenses']
            row['Date'] = datetime.strptime(row['Date'], '%Y-%m-%d')
        except (ValueError, KeyError) as e:
            print(f"Error processing data: {e}")
            return None
    
    return data

def perform_analysis(data):
    """Perform various analyses on the data"""
    if not data:
        return None
    
    # Overall statistics
    total_sales = sum(row['Sales'] for row in data)
    total_expenses = sum(row['Expenses'] for row in data)
    total_profit = total_sales - total_expenses
    profit_margin = (total_profit / total_sales) * 100 if total_sales > 0 else 0
    
    # Group by product
    products = {}
    for row in data:
        product = row['Product']
        if product not in products:
            products[product] = {'sales': 0, 'expenses': 0, 'profit': 0, 'count': 0}
        products[product]['sales'] += row['Sales']
        products[product]['expenses'] += row['Expenses']
        products[product]['profit'] += row['Profit']
        products[product]['count'] += 1
    
    # Group by region
    regions = {}
    for row in data:
        region = row['Region']
        if region not in regions:
            regions[region] = {'sales': 0, 'expenses': 0, 'profit': 0, 'count': 0}
        regions[region]['sales'] += row['Sales']
        regions[region]['expenses'] += row['Expenses']
        regions[region]['profit'] += row['Profit']
        regions[region]['count'] += 1
    
    # Group by month
    months = {}
    for row in data:
        month = row['Date'].strftime('%Y-%m')
        if month not in months:
            months[month] = {'sales': 0, 'expenses': 0, 'profit': 0, 'count': 0}
        months[month]['sales'] += row['Sales']
        months[month]['expenses'] += row['Expenses']
        months[month]['profit'] += row['Profit']
        months[month]['count'] += 1
    
    # Calculate averages
    for product in products:
        products[product]['avg_sales'] = products[product]['sales'] / products[product]['count']
        products[product]['avg_profit'] = products[product]['profit'] / products[product]['count']
    
    for region in regions:
        regions[region]['avg_sales'] = regions[region]['sales'] / regions[region]['count']
        regions[region]['avg_profit'] = regions[region]['profit'] / regions[region]['count']
    
    analysis_results = {
        'overall': {
            'total_sales': total_sales,
            'total_expenses': total_expenses,
            'total_profit': total_profit,
            'profit_margin': profit_margin
        },
        'products': products,
        'regions': regions,
        'months': months
    }
    
    return analysis_results

def create_charts(analysis_results):
    """Create visualizations for the report"""
    products = analysis_results['products']
    months = analysis_results['months']
    
    # Product performance chart
    product_names = list(products.keys())
    product_sales = [products[p]['sales'] for p in product_names]
    product_profits = [products[p]['profit'] for p in product_names]
    
    x = np.arange(len(product_names))
    width = 0.35
    
    plt.figure(figsize=(10, 6))
    plt.bar(x - width/2, product_sales, width, label='Sales')
    plt.bar(x + width/2, product_profits, width, label='Profit')
    plt.xlabel('Products')
    plt.ylabel('Amount ($)')
    plt.title('Sales and Profit by Product')
    plt.xticks(x, product_names)
    plt.legend()
    plt.tight_layout()
    plt.savefig('product_performance.png')
    plt.close()
    
    # Monthly trend chart
    month_names = list(months.keys())
    monthly_sales = [months[m]['sales'] for m in month_names]
    monthly_profits = [months[m]['profit'] for m in month_names]
    
    plt.figure(figsize=(10, 6))
    plt.plot(month_names, monthly_sales, marker='o', label='Sales')
    plt.plot(month_names, monthly_profits, marker='s', label='Profit')
    plt.xlabel('Month')
    plt.ylabel('Amount ($)')
    plt.title('Monthly Sales and Profit Trend')
    plt.legend()
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.savefig('monthly_trend.png')
    plt.close()
    
    print("Visualizations created successfully.")

class PDFReport(FPDF):
    """Custom PDF report class"""
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Sales Performance Report', 0, 1, 'C')
        self.ln(10)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
    
    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(5)
    
    def chapter_body(self, body):
        self.set_font('Arial', '', 12)
        self.multi_cell(0, 10, body)
        self.ln()
    
    def add_table(self, headers, data, col_widths=None):
        self.set_font('Arial', 'B', 12)
        
        if col_widths is None:
            col_width = self.w / (len(headers) + 1)
            col_widths = [col_width] * len(headers)
        
        # Headers
        for i, header in enumerate(headers):
            self.cell(col_widths[i], 10, header, 1)
        self.ln()
        
        # Data
        self.set_font('Arial', '', 12)
        for row in data:
            for i, item in enumerate(row):
                self.cell(col_widths[i], 10, str(item), 1)
            self.ln()

def generate_report(analysis_results, filename='sales_report.pdf'):
    """Generate PDF report with analysis results"""
    pdf = PDFReport()
    pdf.add_page()
    
    # Title
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Sales Performance Analysis Report', 0, 1, 'C')
    pdf.ln(10)
    
    # Date
    pdf.set_font('Arial', 'I', 12)
    pdf.cell(0, 10, f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1, 'C')
    pdf.ln(15)
    
    # Executive Summary
    pdf.chapter_title('Executive Summary')
    overall = analysis_results['overall']
    summary = f"""
    This report provides an analysis of sales performance across different products and regions.
    
    Key Findings:
    - Total Sales: ${overall['total_sales']:,.2f}
    - Total Expenses: ${overall['total_expenses']:,.2f}
    - Total Profit: ${overall['total_profit']:,.2f}
    - Profit Margin: {overall['profit_margin']:.2f}%
    """
    pdf.chapter_body(summary)
    
    # Product Performance
    pdf.add_page()
    pdf.chapter_title('Product Performance')
    
    # Add product performance chart
    pdf.image('product_performance.png', x=10, y=None, w=180)
    pdf.ln(100)
    
    # Product table
    products = analysis_results['products']
    product_table_headers = ['Product', 'Sales', 'Expenses', 'Profit', 'Margin (%)']
    product_table_data = []
    for product, stats in products.items():
        margin = (stats['profit'] / stats['sales']) * 100 if stats['sales'] > 0 else 0
        product_table_data.append([
            product,
            f"${stats['sales']:,.2f}",
            f"${stats['expenses']:,.2f}",
            f"${stats['profit']:,.2f}",
            f"{margin:.2f}%"
        ])
    
    pdf.chapter_title('Product Performance Details')
    pdf.add_table(product_table_headers, product_table_data)
    
    # Regional Performance
    pdf.add_page()
    pdf.chapter_title('Regional Performance')
    
    regions = analysis_results['regions']
    region_table_headers = ['Region', 'Sales', 'Expenses', 'Profit', 'Margin (%)']
    region_table_data = []
    for region, stats in regions.items():
        margin = (stats['profit'] / stats['sales']) * 100 if stats['sales'] > 0 else 0
        region_table_data.append([
            region,
            f"${stats['sales']:,.2f}",
            f"${stats['expenses']:,.2f}",
            f"${stats['profit']:,.2f}",
            f"{margin:.2f}%"
        ])
    
    pdf.add_table(region_table_headers, region_table_data)
    
    # Monthly Trend
    pdf.add_page()
    pdf.chapter_title('Monthly Trend Analysis')
    
    # Add monthly trend chart
    pdf.image('monthly_trend.png', x=10, y=None, w=180)
    pdf.ln(100)
    
    months = analysis_results['months']
    month_table_headers = ['Month', 'Sales', 'Expenses', 'Profit', 'Margin (%)']
    month_table_data = []
    for month, stats in months.items():
        margin = (stats['profit'] / stats['sales']) * 100 if stats['sales'] > 0 else 0
        month_table_data.append([
            month,
            f"${stats['sales']:,.2f}",
            f"${stats['expenses']:,.2f}",
            f"${stats['profit']:,.2f}",
            f"{margin:.2f}%"
        ])
    
    pdf.chapter_title('Monthly Performance Details')
    pdf.add_table(month_table_headers, month_table_data)
    
    # Conclusion
    pdf.add_page()
    pdf.chapter_title('Conclusion and Recommendations')
    
    # Find best performing product
    products = analysis_results['products']
    best_product = max(products.items(), key=lambda x: x[1]['profit'])
    worst_product = min(products.items(), key=lambda x: x[1]['profit'])
    
    # Find best performing region
    regions = analysis_results['regions']
    best_region = max(regions.items(), key=lambda x: x[1]['profit'])
    
    profit_margin = analysis_results['overall']['profit_margin']
    
    conclusion = f"""
    Based on the analysis of the sales data:
    
    1. The best performing product is {best_product[0]} with a profit of ${best_product[1]['profit']:,.2f}.
    2. The product needing improvement is {worst_product[0]} with a profit of ${worst_product[1]['profit']:,.2f}.
    3. The best performing region is {best_region[0]} with a profit of ${best_region[1]['profit']:,.2f}.
    4. The overall profit margin is {profit_margin:.2f}%, which is {'good' if profit_margin > 20 else 'satisfactory' if profit_margin > 10 else 'needs improvement'}.
    
    Recommendations:
    - Focus marketing efforts on {best_product[0]} as it's the most profitable product.
    - Investigate reasons for lower performance of {worst_product[0]} and develop improvement strategies.
    - Expand operations in {best_region[0]} region as it shows the highest profitability.
    - Consider cost reduction strategies if profit margin remains below target.
    """
    
    pdf.chapter_body(conclusion)
    
    # Save the PDF
    pdf.output(filename)
    return filename

def main():
    """Main function to run the report generation process"""
    print("Starting Automated Report Generation...")
    
    # Step 1: Create sample data if needed
    data_filename = 'sales_data.csv'
    create_sample_data(data_filename)
    
    # Step 2: Read and analyze data
    print("Reading and analyzing data...")
    data = read_and_analyze_data(data_filename)
    if not data:
        print("Failed to read or analyze data. Exiting.")
        return
    
    # Step 3: Perform analysis
    analysis_results = perform_analysis(data)
    if not analysis_results:
        print("Failed to perform analysis. Exiting.")
        return
    
    # Display some key statistics
    overall = analysis_results['overall']
    print(f"Total Sales: ${overall['total_sales']:,.2f}")
    print(f"Total Expenses: ${overall['total_expenses']:,.2f}")
    print(f"Total Profit: ${overall['total_profit']:,.2f}")
    print(f"Profit Margin: {overall['profit_margin']:.2f}%")
    
    # Step 4: Create visualizations
    print("Creating visualizations...")
    create_charts(analysis_results)
    
    # Step 5: Generate PDF report
    print("Generating PDF report...")
    report_filename = generate_report(analysis_results)
    print(f"PDF report '{report_filename}' generated successfully!")
    
    # Step 6: Show where files are located
    print("\nGenerated Files:")
    print(f"1. Data file: {os.path.abspath(data_filename)}")
    print(f"2. Product performance chart: {os.path.abspath('product_performance.png')}")
    print(f"3. Monthly trend chart: {os.path.abspath('monthly_trend.png')}")
    print(f"4. PDF report: {os.path.abspath(report_filename)}")
    
    print("\nReport generation completed successfully!")

if __name__ == "__main__":
    main()