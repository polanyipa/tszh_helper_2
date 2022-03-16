import pandas as pd
from tkinter import filedialog


file_name = filedialog.askopenfilename()

#извлечение списка собственников из реестра
data = pd.read_excel(filename)
data = data.drop(data[data['Unnamed: 2'].isna()].index)
data = data.iloc[2:].reset_index()

tenant_list = pd.DataFrame({'name': data['Unnamed: 3'] + ' ' + data['Unnamed: 4'] + ' ' + data['Unnamed: 5'],
                      'flat' : data['Unnamed: 2'],
                      'square': data['Unnamed: 7']})
tenant_list = tenant_list.astype({'square': 'float' })

# создание по шаблона
temp_list = [1, 3, 4, 10, 11, '5 Н']
template = '1010011'
temp = list(template)
result = pd.DataFrame({'flat':temp_list,
                     'присутствовал': 1})
for i in range(len(temp)):
    column_name = 'Вопрос ' + str(i+1)
    result[column_name] = int(temp[i])