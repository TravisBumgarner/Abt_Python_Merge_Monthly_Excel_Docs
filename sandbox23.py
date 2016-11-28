columns_to_keep = [0,1,2]
column_with_unique_id = 0
qty_columns_to_keep = columns_to_keep[:]
qty_columns_to_keep.remove(column_with_unique_id)
qty_columns_to_keep = len(qty_columns_to_keep)
print(qty_columns_to_keep)
