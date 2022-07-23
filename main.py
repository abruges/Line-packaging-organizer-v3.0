import pandas as pd
import datetime as dt
from Info_processor import roasters, orders_dataframe, boxes
from styles import apply_style
import re

# FETCHING DATES:
now = dt.datetime
present_date = now.today().strftime("%d-%m-%Y")

# ORGANIZING ORDERS INFO
orders_info = {}
ordered_products = orders_dataframe['Product/Service'].unique()

for i in ordered_products:
    orders_info[i] = {}

for index, row in orders_dataframe.iterrows():
    product = row['Product/Service']
    order_code = int(row['#'])
    quantity = row['Qty']

    # all_orders dictionary:
    orders_info[product][order_code] = quantity


# CREATING DATAFRAME INDEX BY ORDER NUMBER
order_numbers = orders_dataframe['#'].unique()
order_numbers_list = []

for i in order_numbers:
    order_numbers_list.append(int(i))

order_numbers_list.sort()

fdf = pd.DataFrame({'Orders ⬇️': order_numbers_list})
fdf = fdf.set_index(fdf.columns[0])

# CREATING VALUES FOR DATAFRAME::
dash_list = []

for i in range(len(order_numbers_list)):
    dash_list.append('-')

for i in ordered_products:
    fdf[i] = dash_list

# Adding real data to the table:
for product, info in orders_info.items():
    for order, quantity in info.items():
        fdf.loc[order][product] = int(quantity)

# REORGANIZING BY ROASTER:

# CREATING DATAFRAMES BY ROASTER
df_dictionary = {}
df_length = {}

# IF THE PRODUCT IS FOUND IN THE LIST OF ORDERS, ADD IT TO A NEW LIST
for roaster, products in roasters.items():
    products_found = []
    for product in products:
        if product in ordered_products:
            products_found.append(product)

# IF A MATCH WAS FOUND, THEN CREATE A NEW DATAFRAME WITH THE PRODUCT COLUMNS AND ADD TO DICT:
    if products_found:
        temporary_df = fdf[products_found]

        # REMOVING PARENTHESES FROM COLUMN NAMES
        for column in temporary_df:
            new_name = re.sub("\(.*?\)", "()", column).replace("(", "").replace(")", "")
            temporary_df = temporary_df.rename({column: new_name}, axis=1)

        # ADDING BOX SIZE
        temporary_df.insert(0, 'Box', boxes)

        # ADDING DATAFRAME TO DICTIONARY
        df_dictionary[roaster] = temporary_df
        df_length[roaster] = len(temporary_df.columns)


# WRITING EACH DATAFRAME IN THE FINAL EXCEL:
with pd.ExcelWriter(f'C:/Users/abrug/Desktop/LNV Orders/Logistics/Logistics_{present_date}.xlsx') as writer:
    for name, df in df_dictionary.items():

        df.to_excel(writer, sheet_name=name)


# APPLYING STYLES TO EXCEL
table_height = len(boxes) + 1

red_box_index = [i for i, x in enumerate(boxes) if x == '🧰']
red_box_index = [x+2 for x in red_box_index]

small_box_index = [i for i, x in enumerate(boxes) if x == '📦 18x14x12']
small_box_index = [x+2 for x in small_box_index]

for name, df in df_dictionary.items():
    apply_style(name, table_height, int(df_length[name])+2, red_box_index, small_box_index)