import pandas as pd
from tkinter import filedialog
from openpyxl import load_workbook


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


def decoder(a):
    code = {'0': '001',
            '1': '010',
            '2': '100'}
    result = ''.join(map(lambda x: code[x], list(a)))
    return list(map(int, result))


def template_answer(n):
    while True:
        temp = input('Шаблон ответа?')
        if (len(temp) != n) | (temp.strip('012') != ''):
            print('некорректный ввод')
        else:
            return temp


def template(answer, columns):
    # запрос списка квартир для шаблона, формирование шаблона
    tmp_list = input('Список квартир(через запятую)?')
    result = pd.DataFrame(columns=columns)
    result.flat = list(tmp_list.split(','))
    result.iloc[:, 1:5] = [1, 1, 0, 0]
    result.iloc[:, 5:] = decoder(answer)
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
        for i in flats.loc[flats.flat.isin(previous.flat), 'flat'].values:
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
    columns = ['flat', 'filled', '1', '2', '3']
    for i in range(numquest * 3):
        columns.append(str(i + 4))
    result = pd.DataFrame(columns=columns)

    i = 0
    while i == 0:
        answer = template_answer(numquest)
        new_result = template(answer, columns)
        new_result = template_check(new_result, df, result)
        result = pd.concat([result, new_result], axis=0)

        j = 0
        while j == 0:
            confirm = input('ввести еще один шаблон? (y/n)')
            if confirm == 'y':
                j = 1
            elif confirm == 'n':
                i, j = 1, 1
            else:
                print('некорректный ввод')
    df = df.merge(result, how='left', on='flat')
    print('переходим к построчному вводу')
    df = line_enter(df, numquest)
    return df


def line_enter(df, num_question):
    idx_list = df[df['filled'].isna()].index

    i = 0
    while i < idx_list.size:
        answer = input(df.loc[idx_list[i], 'name'])
        if answer == 's':
            i = idx_list.size
            print('ввод прерван')
        elif answer == 'n':
            df.loc[idx_list[i], 'filled'] = 1
            df.loc[idx_list[i], ['1', '2', '3']] = [0, 0, 1]
            i += 1
            print('недействительный бюллетень')
        elif (answer == 'b') & (i == 0):
            pass
        elif (answer == 'b') & (i != 0):
            i -= 1
            print('вернулись на предыдущий шаг')
        elif answer == '':
            df.loc[idx_list[i], 'filled'] = 1
            df.loc[idx_list[i], ['1', '2', '3']] = [0, 1, 0]
            i += 1
        elif (len(answer) != num_question) | (answer.strip('012') != ''):
            print('некорректный ввод')
        else:
            df.loc[idx_list[i], 'filled'] = 1
            df.loc[idx_list[i], ['1', '2', '3']] = [1, 0, 0]
            df.iloc[idx_list[i], 7:] = decoder(answer)
            '''print('\n')
            print(df[df.index == idx_list[i]].to_string())'''
            i += 1
    return df


def single_line_enter(df, numquest):
    columns = ['filled', '1', '2', '3']
    for i in range(numquest * 3):
        columns.append(str(i + 4))
    df = pd.concat([df, pd.DataFrame(columns=columns)], axis=1)
    df = line_enter(df, numquest)
    return df


def result_analyse(df):
    total_square = df.square.sum()
    # Тут можно вставить возможность использовать фиксированную общую площадь

    results_analyse = pd.DataFrame(index=[1, 2], columns=df.columns[3:])
    results_analyse.iloc[0] = df[df.columns[3:]].mul(df['square'], axis=0).sum()
    results_analyse.iloc[1, :3] = results_analyse.iloc[0, :3] / total_square * 100
    results_analyse.iloc[1, 3:] = results_analyse.iloc[0, 3:] / results_analyse.iloc[0, 0] * 100
    results_analyse.iloc[0] = results_analyse.iloc[0].apply(lambda x: round(x, 1))
    results_analyse.iloc[1] = results_analyse.iloc[1].apply(lambda x: str(round(x, 1)) + ' %')
    results_analyse['total'] = [round(total_square, 1), None]
    return results_analyse


def print_result(df1, df2):
    arrays = [['', ' ', 'площадь'],
              ['Собственник', '№ кв.', 'помещения']]
    col1 = pd.MultiIndex.from_arrays(arrays)
    questions = []
    for i in range(int((df1.shape[1] - 6) / 3)):
        questions.append('Вопрос ' + str(i + 1))
    col2 = pd.MultiIndex.from_product([['участие'], ['да', 'нет', 'недейств']])
    col3 = pd.MultiIndex.from_product([questions, ['Да', 'Возд', 'Нет']])
    col4 = pd.MultiIndex.from_arrays([['общая'], ['площадь']])
    out1 = df1.copy(deep=True)
    out2 = df2.copy(deep=True)
    out1.columns = col1.append(col2).append(col3)
    out2.columns = col2.append(col3).append(col4)
    print(out1.to_string(index=False))
    print('\n', '\n')
    print(out2.to_string(index=False))
    return


def safe_to_excel(df):
    # переставим столбцы в нужном порядке
    c = df.columns.tolist()
    col = c[1:2] + c[:1] + c[3:] + c[2:3]
    df = df[col]
    name = r'C:\Users\User_1\Documents\1. Проекты\python\tszh_helper\template\template.xlsx'
    book = load_workbook(name)
    writer = pd.ExcelWriter(name, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df.to_excel(writer, startrow=6, startcol=0, header=False, index=False)
    writer.save()
    return


def tszh_helper():
    member_list, directory = member_list_import()
    num_quest = number_of_q()
    '''i = 0
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
            print('некорректный ввод')'''


    # Временный путь
    output = template_enter(member_list, num_quest)
    # удаляем вспомогательный столбец - флаг обработки строки
    output = output.drop(columns='filled')
    # нужно бы заменить нули на пустые ячейки
    analyse = result_analyse(output)
    print_result(output, analyse)
    safe_to_excel(output)


    '''res, analyse, total_square = result_analyse(output, num_quest)
    safe_to_excel(directory, res, analyse, total_square)

    print('\n', '\n')
    print(res.head(5).to_string())
    print('\n', '\n')
    print(analyse.to_string())
    print('\n', '\n')
    print(total_square.to_string())
    print('Обработка завершена')'''


if __name__ == '__main__':
    tszh_helper()
