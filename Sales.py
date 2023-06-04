# Importing required libraries
import pandas as pd
import matplotlib.pyplot as plt

# Loading data
data = pd.read_excel('Superstore.xls')

# Display the first few rows of the data
print(data.head())

# Basic information about the data
print(data.info())

# Checking for missing values
print(data.isnull().sum())

# Sales trend over the years
data['Order Date'] = pd.to_datetime(data['Order Date'])
data['Year'] = data['Order Date'].dt.year
sales_trend = data.groupby('Year')['Sales'].sum()
sales_trend.plot(kind='line', marker='o')
plt.title('Sales Trend Over the Years')
plt.xlabel('Year')
plt.ylabel('Total Sales')
plt.show()

# Top 10 profitable products
top_products = data.groupby('Product Name')['Profit'].sum().sort_values(ascending=False).head(10)
top_products.plot(kind='bar')
plt.title('Top 10 Profitable Products')
plt.xlabel('Product Name')
plt.ylabel('Total Profit')
plt.xticks(rotation=90)
plt.show()

# Analyzing sales and profit in different regions
region_data = data.groupby('Region')['Sales', 'Profit'].sum()
region_data.plot(kind='bar', subplots=True)
plt.title('Sales and Profit in Different Regions')
plt.xlabel('Region')
plt.xticks(rotation=0)
plt.show()
