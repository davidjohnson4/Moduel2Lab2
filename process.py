from sys import argv, exit
import os
import re
from datetime import date
import pandas as pd

def get_sales_csv():
   
   #checks if the command line parameters were given
    if len(argv) >= 2:
        sals_csv = argv[1]

        #checks if file even exists
        if os.path.isfile(sals_csv):
            return sals_csv
        else:
            print('error: no csv file path provided')
            exit('script is aborted')
    else:
        print('error: no csv file path provided')
        exit('script is aborted')

def get_order_dir(sales_csv):
    
    #get the directory path of sales data
    sales_dir = os.path.dirname(sales_csv)
    #determine orders directory name
    todays_date = date.today().isoformat()
    order_dir_name = 'orders_' + todays_date
    #build the full path of the orders directory
    order_dir = os.path.join(sales_dir, order_dir_name)
    #make the orders
    if not os.path.exists(order_dir):
        os.makedirs(order_dir)
    
    return order_dir

def split_sales_into_orders(sales_csv, order_dir):
    
    sales_df = pd.read_csv(sales_csv)
    sales_df.insert(7, "TOTAL PRICE", sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])

    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)

    for order_id, order_df in sales_df.groupby('ORDER ID'):
        order_df.drop(columns=['ORDER ID'], inplace=True)

        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE' : ['GRAND TOTAL :'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])

        #the save path of the order file
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name)
        order_file_name = 'Order' + str(order_id) + '_' + customer_name + '.xlsx'
        order_file_path = os.path.join(order_dir, order_file_name)

        #saving the order
        sheet_name = 'Order #' + str(order_id)
        
        
        #making it fancy
        writer = pd.ExcelWriter(order_file_path, engine='xlsxwriter')
        order_df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        money_fmt = workbook.add_format({'num_format': '$#,##0'})

        worksheet.set_column('A:I', 20)
        

        writer.save()


sales_csv = get_sales_csv()
order_dir = get_order_dir(sales_csv)
split_sales_into_orders(sales_csv, order_dir)