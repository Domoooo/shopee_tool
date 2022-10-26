import csv
import pandas as pd

# FILE_PATH = r'D:\vs_workplace\shopee_orders\inv_data\JoinTable-Update_20221025.csv'
FILE_PATH = r'D:\vs_workplace\shopee_orders\inv_data\JoinTable-Update_simple.csv'

df = pd.read_csv(FILE_PATH, sep='\t', encoding="utf16")
total_rows = len(df)

inv_lst = [[] for _ in range(total_rows)]

keys_list = ['SKN', 'STR2INV', 'STR8INV', 'STR10INV', 'STR16INV', 'STR18INV', 'STR25INV', 'STR29INV', 'STR40INV', 'STR41INV']

with open(FILE_PATH, 'r', newline='', encoding='utf16') as f:
    rows = csv.DictReader(f, delimiter='\t')

    i = 0
    for row in rows:
        [inv_lst[i].append(int(row[k])) for k in keys_list]
        i += 1

# print(inv_lst)