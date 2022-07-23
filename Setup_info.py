import pandas as pd

# Boxes info
box_capacity_df = pd.read_excel(f'C:/Users/abrug/Desktop/LNV Orders/Table setup.xlsm', sheet_name="Boxes")
box_capacity_df.set_index('Capacity', inplace=True)

# Styling info
styles_df = pd.read_excel(f'C:/Users/abrug/Desktop/LNV Orders/Table setup.xlsm', sheet_name="Color Palette")
styles_df.set_index('Section', inplace=True)

bt_color = str(styles_df.loc['Base table color']['Color code (HEX)'])
tt_color = str(styles_df.loc['Top table color']['Color code (HEX)'])
rb_color = str(styles_df.loc['Red box color']['Color code (HEX)'])
sb_color = str(styles_df.loc['Small box color']['Color code (HEX)'])

#Suppliers
products_raw_data = pd.read_excel('C:/Users/abrug/Desktop/LNV Orders/Table setup.xlsm', sheet_name="Suppliers")
product_supplier_dataframe = pd.DataFrame(products_raw_data)