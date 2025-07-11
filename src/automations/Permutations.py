import pandas as pd
from itertools import permutations

# Replace with your actual Excel file path
input_file = 'index2.xlsx'

# Read only the first column of the first sheet
df = pd.read_excel(input_file, sheet_name=0, usecols=[0])

# Drop NaNs and extract unique values
values = df.iloc[:, 0].dropna().unique()

# Generate all ordered permutations of length 2 as a list
perms = list(permutations(values, 2))

# Prepare a list to collect rows
rows = []

character = ' ('

kg_list = [10, 20, 35, 50, 70, 100, 300, 500, 'KG EXTRA']


for kg in kg_list:
    third = kg
    for first, second in perms:
        # Extract the first three characters of each string
        first_index = first.find(character)
        second_index = second.find(character)
        prefix1 = str(first)[:first_index]
        prefix2 = str(second)[:second_index]

        # Drop if prefixes are identical, or if they are 'GRU' and 'CGH'
        if prefix1 == prefix2 or (prefix1 == 'SÃO PAULO - P.A' and prefix2 == 'SÃO PAULO - MATRIZ') or (prefix1 == 'SÃO PAULO - MATRIZ' and prefix2 == 'SÃO PAULO - P.A'):
            continue

        rows.append((first, second, third))

# Convert to DataFrame and save to CSV
result_df = pd.DataFrame(rows, columns=['first', 'second', 'value'])

result_df.to_excel('permutations.xlsx', index=False)