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


def template_flats(answer):
    tmp_list = input('Список квартир(через запятую)?')
    result = pd.DataFrame({'flat': list(tmp_list.split(',')),
                     'присутствовал': 1})

    for i in range(len(answer)):
        column_name = 'Вопрос ' + str(i + 1)
        result[column_name] = int(answer[i])
    return result


def temp_flat_list_check(flats, member_list, previous):
    return flats


def template_enter(df, numquest):
    columns = ['flat', 'присутствовал']
    for i in range(numquest):
        columns.append('Вопрос ' + str(i + 1))
    result = pd.DataFrame(columns=columns)

    i = 0
    while i == 0:
        answer = template_list(numquest)
        new_result = template_flats(answer)
        new_result = temp_flat_list_check(new_result, df, result)
        result = pd.concat([result, new_result], axis=0)

        j = 0
        while j == 0:
            confirm = input('ввести еще один шаблон? (y/n)')
            if confirm == 'y':
                j = 1
            elif confirm == 'n':
                j = 1
                i = 1
            else:
                print('некорректный воод')
    df = df.merge(result, how='left', on='flat')
    df = line_enter(df, numquest)
    return df


def line_enter(df, num_question):
    idx_list = df[df['присутствовал'].isna()].index

    i = 0
    while i < idx_list.size:
        answer = input(df.loc[idx_list[i], 'name'])
        if answer == 's':
            i = idx_list.size
            print('ввод прерван')
        if (answer == 'b') & (i == 0):
            pass
        elif (answer == 'b') & (i != 0):
            i -= 1
            print('вернулись на предыдущий шаг')
        elif answer == '':
            df.loc[idx_list[i], 'присутствовал'] = 0
            i += 1
        elif (len(answer) != num_question) | (answer.strip('01') != ''):
            print('некорректный ввод')
        else:
            df.loc[idx_list[i], 'присутствовал'] = 1
            df.iloc[idx_list[i], 4:] = list(map(int, list(answer)))
            print('\n')
            print(df[df.index == idx_list[i]].to_string())
            i += 1
    return df


def single_line_enter(df, numquest):
    columns = ['присутствовал']
    for i in range(numquest):
        columns.append('Вопрос ' + str(i + 1))
    df = pd.concat([df, pd.DataFrame(columns=columns)], axis=1)
    df = line_enter(df, numquest)
    return df


def tszh_helper():
    member_list = member_list_import()
    num_quest = number_of_q()
    i = 0
    while i == 0:
        ans = input('Ввести шаблон? (y/n)')
        if ans == 'y':
            print('задаем шаблон')
            output = template_enter(member_list, num_quest)
            i += 1
        elif ans == 'n':
            print('переходим к построчному вводу')
            output = single_line_enter(member_list, num_quest)
            i += 1
        else:
            print('некорректный ввод')

    print(output.head(10).to_string())


if __name__ == '__main__':
    tszh_helper()
