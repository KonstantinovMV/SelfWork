import pandas as pd
import os.path

LastNamesList = ['ов', 'ова', 'ин', 'ина', 'ев', 'ева', 'ын', 'ына', 'ский', 'ская', 'цкий', 'цкая', 'ый', 'ая', 'ой',
                 'ий']
SurNamesList = ['ович', 'овна', 'евич', 'евна', 'ич', 'ична', 'инична']
ExceptList = ["Агриппина", "Аделина", "Акулина", "Алевтина", "Алина", "Альбина", "Альвина", "Амина", "Ангелина",
              "Антонина", "Арина", "Валентина", "Василина", "Веселина", "Галина", "Георгина", "Гражина", "Дарина",
              "Дина", "Евангелина", "Екатерина", "Жозефина", "Зарина", "Ирина", "Капитолина", "Карина", "Каролина",
              "Катарина", "Климентина", "Кристина", "Лина", "Магдалина", "Мадина", "Мальвина", "Марина", "Марселина",
              "Михайлина", "Нина", "Полина", "Регина", "Руфина", "Сабина", "Сабрина", "Северина", "Тина", "Устина",
              "Фаина", "Христина", "Эвелина", "Элина", "Эрнестина", "Ярина"]


def number_checker(number: str) -> tuple:
    """Проверяет номер на предмет корректности, заменяет первую цифрц 8 на +7.
    Возвращает корретный номер в формате +7 или ошибку"""
    if number.startswith("+7") and len(number) == 12:
        return number, None
    elif number.startswith("8") and len(number) == 11:
        return '+7' + number[1:], None
    else:
        return number, 'incorrect_number'


def sorter_names(item: str) -> str:
    """Проверяет слова на предмет схожести с фамилией и отчеством,
    возвращает отсортированный список и ошибку если она есть"""
    for LN in LastNamesList:
        if item.endswith(LN) and item not in ExceptList:
            return 'LastName'
    for SN in SurNamesList:
        if item.endswith(SN):
            return 'SurName'


def SortToExcelFormat(data: dict) -> dict:
    """Проходит по идексам словаря и создает словарь в формате, удобном для записи в экзель через Pandas"""
    Resultdata = {'Numbers': [], 'LastNames': [], 'Names': [], 'SurNames': [], 'Errors': []}

    for x in data.values():
        Resultdata['Numbers'].append(x['Number'])
        Resultdata['LastNames'].append(x['LastName'])
        Resultdata['Names'].append(x['Name'])
        Resultdata['SurNames'].append(x['SurName'])
        Resultdata['Errors'].append(x['Errors_name'])

    return Resultdata


if os.path.exists(r'C:\Users\Макс\PycharmProjects\SimpleProjects\NameList\Nameslist.txt'):
    if os.stat(r'C:\Users\Макс\PycharmProjects\SimpleProjects\NameList\Nameslist.txt').st_size == 0:
        print('Error: file is empty')
    else:
        MyDataFrame = {}
        names_list = open('NamesList.txt', 'r', encoding="utf-8")
        id_count = 1

        for line in names_list:
            MyDataFrame[id_count] = {'Number': None, 'LastName': None, 'Name': None, 'SurName': None, 'Errors_name': None}
            for word in line.split():
                if word.isdigit() or word.startswith('+7'):
                    CorrectNumber, Error = number_checker(word)
                    MyDataFrame[id_count]['Number'] = CorrectNumber
                    if Error:
                        MyDataFrame[id_count]['Errors_name'] = Error

                else:
                    word = word.lower().title()
                    NameType = sorter_names(word)
                    if NameType == "LastName":
                        MyDataFrame[id_count]['LastName'] = word
                    elif NameType == "SurName":
                        MyDataFrame[id_count]['SurName'] = word
                    else:
                        MyDataFrame[id_count]['Name'] = word
            id_count += 1

        SortedDataFrame = SortToExcelFormat(MyDataFrame)

        SortedResult = pd.DataFrame(SortedDataFrame)
        SortedResult.to_excel('./sorted_names.xlsx', sheet_name='Contacts', index=False)
        names_list.close()

else:
    print('Error: no file detected')
