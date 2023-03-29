#!/usr/bin/env python
# coding: utf-8

# Created on 3/29/2023
# Created by: Christian Daniel Martinez
# Email: Cdmartinez1997@gmail.com
#


import pandas as pd
import tkinter as tk
import datetime
from tkinter import *
from tkinter import filedialog
from tkinter import simpledialog


# Create the GUI to prompt for file selection and save location
root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(title="Select XLSX File", filetypes=[("Excel files", "*.xlsx")])

# Read the XLSX file using pandas
df = pd.read_excel(file_path)

# Create a new dataframe with all the required columns and filled with blank values
new_df = pd.DataFrame(columns=['Customer ID', 'Sales Order/Proposal #', 'Date', 'Ship By', 'Proposal', 'Proposal Accepted',
                               'Closed', 'Quote #', 'Drop Ship', 'Ship to Name', 'Ship to Address-Line One',
                               'Ship to Address-Line Two', 'Ship to City', 'Ship to State', 'Ship to Zipcode',
                               'Ship to Country', 'Customer PO', 'Ship Via', 'Discount Amount', 'Displayed Terms',
                               'Sales Representative ID', 'Accounts Receivable Account', 'Sales Tax ID', 'Invoice Note',
                               'Note Prints After Line Items', 'Statement Note', 'Stmt Note Prints Before Ref',
                               'Internal Note', 'Number of Distributions', 'SO/Proposal Distribution', 'Quantity',
                               'Item ID', 'Description', 'G/L Account', 'Unit Price', 'Tax Type', 'UPC / SKU',
                               'Weight', 'U/M ID', 'U/M No. of Stocking Units', 'Stocking Quantity', 'Stocking Unit Price',
                               'Amount', 'Job ID', 'Sales Tax Agency ID'])

save_path = filedialog.asksaveasfilename(title="Save CSV File", defaultextension=".csv", filetypes=[("CSV files", "*.csv")])

now = datetime.datetime.now()
date = now.strftime("%m/%d/%Y")

# Fill the required columns with data from the original dataframe
new_df['Quantity'] = df['Qty']
new_df['Item ID'] = ''

# Fill empty ItemNo with the value from the row above it
df['ItemNo'].fillna(method='ffill', inplace=True)
df['Unit'].fillna(method='ffill', inplace=True)

# If there is no ItemNo for a row, fill it with the value from the row above it
df['ItemNo'].fillna(method='ffill', inplace=True)
# If there is no Model variable, use Spec as description
df.loc[df['Model'].isna(), 'Description'] = df['Spec'].astype(str)

# Create the Description column using the specified format
new_df['Description'] = df['Description'].fillna(df['Unit'].str.upper() + ". " + "Item #" + df['ItemNo'] + ", Model #" + df['Model'].astype(str) + ", " + df['Spec'].astype(str))



new_df['Date'] = date
new_df['Ship By'] = date
new_df['Proposal'] = 'FALSE'
new_df['Proposal Accepted'] = 'FALSE'
new_df['Closed'] = 'FALSE'
new_df['Drop Ship'] = 'FALSE'
new_df['Ship Via'] = 'LEAST EXPENSIVE'
new_df['Discount Amount'] = '0'
new_df['Displayed Terms'] = 'Prepaid'
new_df['Sales Representative ID'] = 'HEICO'
new_df['Accounts Receivable Account'] = '12000'
new_df['Note Prints After Line Items'] = 'FALSE'
new_df['Stmt Note Prints Before Ref'] = 'FALSE'

new_df['Number of Distributions'] = len(df)
new_df['SO/Proposal Distribution'] = range(1, len(df)+1)


new_df['G/L Account'] = '40100'
new_df['Unit Price'] = df['Sell']
new_df['Tax Type'] = '1'
new_df['Weight'] = '0'

new_df['U/M No. of Stocking Units'] = '1'
new_df['Stocking Quantity'] = new_df['Quantity']
new_df['Stocking Unit Price'] = new_df['Unit Price']
new_df['Amount'] = '-' + df['SellTotal'].astype(str)



while True:
    customer_id = simpledialog.askstring(title="Customer ID", prompt="Enter Customer ID (3 letters followed by 3 numbers):").upper()
    if len(customer_id) == 6 and customer_id[:3].isalpha() and customer_id[3:].isdigit():
        new_df['Customer ID'] = customer_id
        break
while True:
    SO = simpledialog.askstring(title="Sales Order/Proposal #", prompt="Enter Sales Order Number (20 Characters or less):").upper()
    if len(SO) <= 20:
        new_df['Sales Order/Proposal #'] = SO
        break









# Save the new dataframe as a CSV file
new_df.to_csv(save_path, index=False)


# In[ ]:




