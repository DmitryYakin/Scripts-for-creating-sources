import pandas as pd
import json

# set to False for excel format
# set to True for md format
md = False                                   
# date range of report, used in excel format
minDate, maxDate = '28.07.25', '03.08.25'
# name of the sheet for excel format, used in excel format
sheetName =  'Отчет'

# header for md format, used in md format
HEADER = 'Отчет по запросам врачей Астраханской области' 

# list of mo names
MO = pd.read_csv("momar.csv")
# list of requests
requests = pd.read_csv('mar.csv')

requests['doctor_speciality'] = requests['doctor_speciality'].map(lambda x: json.loads(x)[0]['speciality_name']) 
requests['doctor_speciality'] = requests['doctor_speciality'].map(lambda x: x if x else 'Не указана специальность') 

MODict = {}

for i, row in MO.iterrows():
    MODict[row['id']] = row['name']

requests["hospital_name"] = requests['hospital_id'].map(lambda x: MODict[x] if x in MODict else 'Мо не указано')
# print(requests)

if md:
    reportik = [f'# {HEADER}', '']

    for mo in requests['hospital_name'].unique():
        if mo:
            requestsMO = requests[requests['hospital_name'] == mo]
            for speciality in requestsMO['doctor_speciality'].unique():
                if True:
                    reportik.append(f"    - {speciality}: {len(requestsMO[requestsMO['doctor_speciality'] == speciality])}")

            reportik.append('')

    with open('report.md', 'w') as f:
        f.write('\n'.join(reportik))

else:
    reportik = []
    for mo in requests['hospital_name'].unique():
        if mo: 
            requestsMO = requests[requests['hospital_name'] == mo]
            for speciality in requestsMO['doctor_speciality'].unique():
                reportik.append({'МО': mo,
                 'Направление': speciality,
                 'Количество': len(requestsMO[requestsMO['doctor_speciality'] == speciality]),
                 'Период': f'с {minDate} по {maxDate} год'})

    reportik = pd.DataFrame(reportik)
    with pd.ExcelWriter('Отчёт по Марий Эл.xlsx') as w:
        reportik.to_excel(w, sheet_name = sheetName, index = False)

        for col in reportik:
            colLen = max(reportik[col].astype(str).map(len).max(), len(col)) + 2
            colI = reportik.columns.get_loc(col)
            w.sheets[sheetName].set_column(colI, colI, colLen)
