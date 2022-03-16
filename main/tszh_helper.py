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
    members = members.astype({'square': 'float',
                              'flat': 'str',
                              'name': 'str'})
    print(members.head(2).to_string(), '\n', members.tail(2).to_string())
    return members


def number_of_q():
    while True:
        try:
            val = int(input('Количество вопросов?'))
            if val > 0:
                print(val, ' вопросов')
                return val
            else:
                print('некорректный ввод')
        except ValueError:
            print('некорректный ввод')


def template_list(n):
    while True:
        temp = input('Шаблон ответа?')
        if (len(temp) != n) | (temp.strip('01') != ''):
            print('некорректный ввод')
        else:
            return list(map(int, list(temp)))


def template_enter(numquest):
    columns = ['flat', 'присутствовал']
    for i in range(numquest):
        columns.append('Вопрос ' + str(i + 1))
    new_result = pd.DataFrame(columns=columns)
    # это временный вывод
    print(new_result.to_string())


if __name__ == '__main__':
    member_list = member_list_import()
    num_quest = number_of_q()
    i = 0
    while i == 0:
        ans = input('Ввести шаблон? (y/n)')
        if ans == 'y':
            print('задаем шаблон')
            template_enter(num_quest)
            i += 1
        elif ans == 'n':
            print('переходим к построчному вводу')
            i += 1
        else:
            print('некорректный ввод')
