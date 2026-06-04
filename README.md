# Decodelabs-TASK-2

import pandas as pd

# Load the Excel file to understand its structure
file_path = 'Dataset for Data Analytics (1).xlsx'
try:
    xls = pd.ExcelFile(file_path)
    print("Sheet names:", xls.sheet_names)
    
    # Read the first sheet
    df = pd.read_excel(file_path, sheet_name=xls.sheet_names[0])
    print("\nData Info:")
    print(df.info())
    print("\nFirst few rows:")
    print(df.head())
except Exception as e:
    print("Error:", e)

import pandas as pd
import numpy as np

# Load original data
file_path = 'Dataset for Data Analytics (1).xlsx'
df = pd.read_excel(file_path)

# 1. Basic Statistics
numeric_cols = ['Quantity', 'UnitPrice', 'ItemsInCart', 'TotalPrice']
desc_stats = df[numeric_cols].describe().T
desc_stats['median'] = df[numeric_cols].median()
desc_stats = desc_stats[['count', 'mean', 'median', 'min', 'max', 'std']]

# 2. Identify Trends (Aggregations)
# Monthly trend
df['Month_Year'] = df['Date'].dt.to_period('M')
monthly_sales = df.groupby('Month_Year')['TotalPrice'].sum().reset_index()
monthly_sales['Month_Year'] = monthly_sales['Month_Year'].astype(str)

# Product trend
product_sales = df.groupby('Product').agg({'Quantity': 'sum', 'TotalPrice': 'sum'}).reset_index().sort_values(by='TotalPrice', ascending=False)

# 3. Identify Outliers (using IQR method for TotalPrice)
Q1 = df['TotalPrice'].quantile(0.25)
Q3 = df['TotalPrice'].quantile(0.75)
IQR = Q3 - Q1
lower_bound = Q1 - 1.5 * IQR
upper_bound = Q3 + 1.5 * IQR
outliers = df[(df['TotalPrice'] < lower_bound) | (df['TotalPrice'] > upper_bound)][['OrderID', 'Product', 'Quantity', 'UnitPrice', 'TotalPrice']]

# Create an Excel writer object
output_file = 'EDA_Completed_Task.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Write original data
    df.drop(columns=['Month_Year']).to_excel(writer, sheet_name='Original Data', index=False)
    
    # Write basic stats
    desc_stats.to_excel(writer, sheet_name='Basic Statistics')
    
    # Write product sales (trends)
    product_sales.to_excel(writer, sheet_name='Product Trends', index=False)
    
    # Write Monthly sales
    monthly_sales.to_excel(writer, sheet_name='Monthly Sales Trend', index=False)
    
    # Write Outliers
    outliers.to_excel(writer, sheet_name='Outliers Analysis', index=False)

print(f"File saved to {output_file}")
print("\nDescriptive Stats:")
print(desc_stats)
print(f"\nNumber of outliers found: {len(outliers)}")


print("Product Sales Summary:")
print(product_sales.head())

print("\nPayment Method Distribution:")
print(df['PaymentMethod'].value_counts())

print("\nOrder Status Distribution:")
print(df['OrderStatus'].value_counts())
