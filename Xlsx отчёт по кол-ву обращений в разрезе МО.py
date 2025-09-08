import pandas as pd

HEADER = 'Обращения по Республике Удмуртия за июнь 2025 год'

# Read data from CSV files
MO = pd.read_csv("mo.csv")
requests = pd.read_csv('udmall.csv')

# Create a mapping dictionary from MO DataFrame
MODict = dict(zip(MO['id'], MO['name']))

# Map 'hospital_id' to 'hospital_name' using MODict
requests["hospital_name"] = requests['hospital_id'].map(MODict).fillna('МО не указано')

# Group requests by 'hospital_name' and sum the 'count' column
grouped_requests = requests.groupby('hospital_name')['count'].sum().reset_index()

# Initialize the report list with a header for Markdown
reportik = [f'# {HEADER}', '']

# Iterate over each group and append to the report
for index, row in grouped_requests.iterrows():
    mo = row['hospital_name']
    total_count = row['count']
    reportik.append(f'-   {mo}: {total_count}')
    reportik.append('')  # Empty line for separation

# Write the Markdown report
with open('reportpi.md', 'w', encoding='utf-8') as f:
    f.write('\n'.join(reportik))

# Save to Excel
grouped_requests.columns = ['МО', 'Количество']
grouped_requests.to_excel('Республика Удмуртия.xlsx', index=False)
