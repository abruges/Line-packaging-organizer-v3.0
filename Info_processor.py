import pandas as pd
import datetime as dt
from Setup_info import box_capacity_df, product_supplier_dataframe

# FETCHING PRODUCT INVENTORY AND ORDERS:
now = dt.datetime
present_date = now.today().strftime("%d-%m-%Y")

orders_raw_data = pd.read_excel(f'Orders_{present_date}.xlsx', engine="openpyxl")
orders_dataframe = pd.DataFrame(orders_raw_data)
orders_dataframe = orders_dataframe[orders_dataframe['Product/Service'].notna()]

# COUNTING AMOUNT OF PRODUCTS BY ORDER
orders_by_size = {}

for i in orders_dataframe['#'].unique():
    orders_by_size[int(i)] = 0

for index, row in orders_dataframe.iterrows():
    order_number = int(row['#'])
    amount = int(row['Qty'])
    orders_by_size[order_number] += amount

ordered_orders_by_size = []
for order, amount in orders_by_size.items():
    ordered_orders_by_size.append(order)


# FINDING BOX FOR EACH ORDER
boxes = []
x_amounts = []

for i in range(len(ordered_orders_by_size)):
    x = orders_by_size[min(ordered_orders_by_size)]
    x_amounts.append(x)
    ordered_orders_by_size.remove(min(ordered_orders_by_size))


for i in x_amounts:
    capacity_boxes = list(box_capacity_df.index.unique().values)
    box_found = False
    while not box_found:
        x = int(min(capacity_boxes))
        if i <= x:
            boxes.append(str(box_capacity_df.loc[x]['Box size']))
            box_found = True
        else:
            capacity_boxes.remove(x)


# GROUPING PRODUCTS BY ROASTER
roasters_list = product_supplier_dataframe['Preferred Supplier'].unique()
roasters = {}

for i in roasters_list:
    roasters[i] = []

for index, row in product_supplier_dataframe.iterrows():
    roaster = row['Preferred Supplier']
    product = row['Product/Service']

    if product in roasters[roaster]:
        pass
    else:
        roasters[roaster].append(product)
