import pandas as pd
import json

HEADER = 'Отчет по запросам врачей Рязани с 01.01 по 28.01'

MO = pd.read_csv("mo.csv")
requests = pd.read_csv('hack.csv')

requests['doctor_speciality'] = requests['doctor_speciality'].map(lambda x: json.loads(x)[0]['speciality_name']) 
requests['doctor_speciality'] = requests['doctor_speciality'].map(lambda x: x if x else 'Не указана специальность') 

MODict = {}

for i, row in MO.iterrows():
    MODict[row['id']] = row['name']

requests["hospital_name"] = requests['hospital_id'].map(lambda x: MODict[x] if x in MODict else 'Мо не указано')
# print(requests)

reportik = [f'# {HEADER}', '']

for mo in requests['hospital_name'].unique():
    if mo:
        requestsMO = requests[requests['hospital_name'] == mo]
        # requestsMO = requests.query('hospital_name == @mo and doctor_speciality != None')
        reportik.append(f'1. {mo}: {len(requestsMO)}')
        # print(requestsMO)
        for speciality in requestsMO['doctor_speciality'].unique():
            # print(speciality, type(speciality))
            if True:
                reportik.append(f"    - {speciality}: {len(requestsMO[requestsMO['doctor_speciality'] == speciality])}")

        reportik.append('')

with open('report.md', 'w') as f:
    f.write('\n'.join(reportik))