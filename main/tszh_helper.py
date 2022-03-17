import pandas as pd
from tkinter import filedialog
import numpy as np


def member_list_import():
    # чтение файла
    filename = filedialog.askopenfilename()
    directory = filename.rsplit('/', 1)[0]
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
    return members, directory


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


def template_answer(n):
    while True:
        temp = input('Шаблон ответа?')
        if (len(temp) != n) | (temp.strip('012') != ''):
            print('некорректный ввод')
        else:
            return list(map(int, list(temp)))


def template(answer):
    # запрос списка квартир для шаблона, формирование шаблона
    tmp_list = input('Список квартир(через запятую)?')
    result = pd.DataFrame({'flat': list(tmp_list.split(',')),
                     'присутствовал': 1})

    for i in range(len(answer)):
        column_name = 'Вопрос ' + str(i + 1)
        result[column_name] = int(answer[i])
    return result


def template_check(flats, member_list, previous):
    # проверка шаблона на правильность

    # Проверка на то, имеются ли введенные квартиры в реестре
    if not (flats.flat.isin(member_list.flat).all()):
        print('в реестре отсутствуют квартиры:')
        for i in flats.loc[flats.flat.isin(member_list.flat) == False, 'flat'].values:
            print(i)
        print('они не будут учтены в голосовании:')
        flats_out = flats[flats.flat.isin(member_list.flat)]
        flats = flats_out

    # Проверка на то, не были ли введены данные квартиры на предыдущем этапе
    if flats.flat.isin(previous.flat).any():
        print('Эти квартиры были введены на предыдущем этапе:')
        for i in flats.loc[flats.flat.isin(previous.flat) == True, 'flat'].values:
            print(i)
        print('они не будут учтены в голосовании:')
        flats_out = flats[flats.flat.isin(previous.flat) == False]
        flats = flats_out

    # Проверка на то, не введены ли квартиры дважды
    if flats.flat.duplicated().any():
        print('Номер квартиры введен дважды:')
        for i in flats.loc[flats.flat.duplicated(), 'flat'].values:
            print(i)
        print('Проверьте правильность ввода')
        flats = flats.drop_duplicates()

    # Подтверждение принятия шаблона
    i = 0
    while i == 0:
        confirm = input('Принять шаблон? (y/n)')
        if confirm == 'y':
            i = 1
        elif confirm == 'n':
            flats = flats.iloc[0:0]
            i = 1
        else:
            print('некорректный ввод')

    return flats


def template_enter(df, numquest):
    columns = ['flat', 'присутствовал']
    for i in range(numquest):
        columns.append('Вопрос ' + str(i + 1))
    result = pd.DataFrame(columns=columns)

    i = 0
    while i == 0:
        answer = template_answer(numquest)
        new_result = template(answer)
        new_result = template_check(new_result, df, result)
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
                print('некорректный ввод')
    df = df.merge(result, how='left', on='flat')
    print('переходим к построчному вводу')
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
        elif answer == 'n':
            df.loc[idx_list[i], 'присутствовал'] = 2
            i += 1
            print('недействительный бюллетень')
        elif (answer == 'b') & (i == 0):
            pass
        elif (answer == 'b') & (i != 0):
            i -= 1
            print('вернулись на предыдущий шаг')
        elif answer == '':
            df.loc[idx_list[i], 'присутствовал'] = 0
            i += 1
        elif (len(answer) != num_question) | (answer.strip('012') != ''):
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


def result_analyse(df, numquest):
    results = pd.DataFrame(columns=pd.MultiIndex.from_arrays([['кв №', 'Собственник'], [' ', ' ']]))
    results[[('кв №', ' '), ('Собственник', ' ')]] = df[['flat', 'name']]
    column_name = ['участие']
    for i in range(numquest):
        name = 'Вопрос ' + str(i + 1)
        column_name.append(name)
    multi_col = pd.MultiIndex.from_product([column_name, ['Да', 'Нет', 'Воздерж']],
                                   names=['1', '2'])
    result_2 = pd.DataFrame(columns=multi_col)
    results = pd.concat([results, result_2], axis=0)

    results[('участие', 'Да')] = np.where(df['присутствовал'] == 1, 1, None)
    results[('участие', 'Нет')] = np.where(df['присутствовал'] == 0, 1, None)
    results[('участие', 'Воздерж')] = np.where(df['присутствовал'] == 2, 1, None)

    for col in column_name[1:]:
        results[(col, 'Да')] = np.where(df[col] == 2, 1, None)
        results[(col, 'Нет')] = np.where(df[col] == 0, 1, None)
        results[(col, 'Воздерж')] = np.where(df[col] == 1, 1, None)


    for col in column_name[1:]:
        results[(col, 'Да')] = np.where(df[col] == 2, 1, None)
        results[(col, 'Нет')] = np.where(df[col] == 0, 1, None)
        results[(col, 'Воздерж')] = np.where(df[col] == 1, 1, None)

    results_analyse = pd.DataFrame(columns=multi_col, index=[0, 1])

    total_square = df.square.sum()
    enable_square = df[df['присутствовал'] == 1]['square'].sum()

    results_analyse.loc[0, ('участие', 'Да')] = df[df['присутствовал'] == 1]['square'].sum()
    results_analyse.loc[0, ('участие', 'Нет')] = df[df['присутствовал'] == 0]['square'].sum()
    results_analyse.loc[0, ('участие', 'Воздерж')] = df[df['присутствовал'] == 2]['square'].sum()
    results_analyse.loc[1, ('участие', 'Да')] = (str(round(df[df['присутствовал'] == 1]['square'].sum()
                                                           / total_square * 100, 1))+' %')
    results_analyse.loc[1, ('участие', 'Нет')] = (str(round(df[df['присутствовал'] == 0]['square'].sum()
                                                           / total_square * 100, 1))+' %')
    results_analyse.loc[1, ('участие', 'Воздерж')] = (str(round(df[df['присутствовал'] == 2]['square'].sum()
                                                           / total_square * 100, 1))+' %')
    for col in column_name[1:]:
        results_analyse.loc[0, (col, 'Да')] = df[df[col] == 2]['square'].sum()
        results_analyse.loc[0, (col, 'Нет')] = df[df[col] == 0]['square'].sum()
        results_analyse.loc[0, (col, 'Воздерж')] = df[df[col] == 1]['square'].sum()
        results_analyse.loc[1, (col, 'Да')] = (str(round(df[df[col] == 2]['square'].sum()
                                                           / enable_square * 100, 1))+' %')
        results_analyse.loc[1, (col, 'Нет')] = (str(round(df[df[col] == 0]['square'].sum()
                                                           / enable_square * 100, 1))+' %')
        results_analyse.loc[1, (col, 'Воздерж')] = (str(round(df[df[col] == 1]['square'].sum()
                                                           / enable_square * 100, 1))+' %')

    total = pd.DataFrame({'общая площадь': total_square}, index=[0])

    return results, results_analyse, total


def safe_to_excel(directory, df1, df2, df3):
    name = directory + '\\' + 'out.xls'
    with pd.ExcelWriter(name, engine='xlsxwriter') as writer2:
        df1.to_excel(writer2, sheet_name='Spec', startrow=0, startcol=0, index=True, header=True)
        df3.to_excel(writer2, sheet_name='Spec', startrow=df1.index.size + 5, startcol=0, index=True, header=True)
        df2.to_excel(writer2, sheet_name='Spec', startrow=df1.index.size + 5, startcol=2, index=True, header=True)
    return


def tszh_helper():
    member_list, directory = member_list_import()
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

    res, analyse, total_square = result_analyse(output, num_quest)
    safe_to_excel(directory, res, analyse, total_square)

    print('\n', '\n')
    print(res.head(5).to_string())
    print('\n', '\n')
    print(analyse.to_string())
    print('\n', '\n')
    print(total_square.to_string())
    print('Обработка завершена')


if __name__ == '__main__':
    tszh_helper()
