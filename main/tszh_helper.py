import pandas as pd
from tkinter import filedialog


def member_list_import():
    # чтение файла
    filename = filedialog.askopenfilename()
    data = pd.read_excel(filename)

    # извленчение нужных данных (имена, номера квартир, площадь)
    data = data.drop(data[data['Unnamed: 2'].isna()].index)
    data = data.iloc[2:].reset_index()
    members = pd.DataFrame(
        {'name': data['Unnamed: 3'] + ' ' + data['Unnamed: 4'] + ' ' + data['Unnamed: 5'],
            'flat': data['Unnamed: 2'],
            'square': data['Unnamed: 7']})
    members = members.astype({'square': 'float'})
    return members



'''
# создание по шаблона
temp_list = [1, 3, 4, 10, 11, '5 Н']
template = '1010011'
temp = list(template)
result = pd.DataFrame({'flat':temp_list,
                     'присутствовал': 1})
for i in range(len(temp)):
    column_name = 'Вопрос ' + str(i+1)
    result[column_name] = int(temp[i])'''

if __name__ == '__main__':
    member_list = member_list_import()
    print(member_list.head(2).to_string(), '\n', member_list.tail(2).to_string())
