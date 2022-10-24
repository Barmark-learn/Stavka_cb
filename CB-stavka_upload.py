import pandas as pd
import datetime as dt


# Функция перевода в нормальную дату для экселя
def date_normal(nr):
    for i in range(len(nr)):
        nr[i] = nr[i].strftime('%d.%m.%Y ')


# Создаём датафрейм
file_df = pd.read_excel('россия-ключевая-ставка-цб.xlsx')

# Имя столбца с датами и ставкой и колличество строк
name_colum = file_df.columns[0]
name_colum_dop = file_df.columns[1]
count_line = file_df.shape[0]
n = dt.datetime.today().date()

# Переводим датафрейм в словарь
dict_file = file_df.to_dict()

# Переводим в список значения ставки
rate_list = file_df[name_colum_dop].tolist()

# Определяем сегодняшнюю дату и добавляем ее в исходный дф
now_date = [dt.datetime.today().date()]

date_normal(now_date)
now_date.append(rate_list[0])

file_df.loc[-1] = now_date  # adding a row
file_df.index = file_df.index + 1  # shifting index
file_df = file_df.sort_index()  # sorting by index

# Создаём список для ставок и дат
lst_r = []
lst = []

# Исходные даты в виде списка
date_list = file_df[name_colum].tolist()


# Цикл прохода по всем строчкам дат и нахождения промежутков между датами,
# а также заполнение ставками в соответствии с датами
for i in range(count_line - 1):
    if dict_file[name_colum][i] != dict_file[name_colum][i + 1]:
        lst_1 = list(pd.date_range(file_df[name_colum][i + 1], file_df[name_colum][i]))
        lst_2 = [rate_list[i + 1]] * len(lst_1)
        lst_r += lst_2[1:]
        lst_1.reverse()
        lst += lst_1[1:]
        lst_1.clear()
        lst_2.clear()
lst.insert(0, file_df[name_colum][0])
lst_r.insert(0, rate_list[0])
# Переводим в нормальную дату (тип строка)
date_normal(lst)

# Заполняем новый датафрейм датами и ставкаи
df = pd.DataFrame({name_colum: lst, name_colum_dop: lst_r})

# Переводим наш датафрейм в эксель
df.to_excel('Книга1.xlsx', sheet_name='Advanced', index=False)
