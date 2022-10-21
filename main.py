# Author Loik Andrey 7034@balancedv.ru
import pandas as pd
import os
from loguru import logger


"""logger.add(config.FILE_NAME_CONFIG,
           format="{time:DD/MM/YY HH:mm:ss} - {file} - {level} - {message}",
           level="INFO",
           rotation="1 month",
           compression="zip")"""


def search_file():
    for item in os.listdir():
        if item.endswith('xls'):
            logger.info(f"Find file 'xls': {item}")
            return item
    logger.error(f"Can't find file 'xls'")

def read_file(file):
    return pd.read_excel(file)


def rebuild_df(df):
    mask_end = df == 'Итого за вес'
    end_row = df[mask_end].dropna(axis=0, how='all').index.values
    mask_start = df == 'Марка'
    start_row = df[mask_start].dropna(axis=0, how='all').index.values

    df_head = df.iloc[0:2].dropna(axis=0, how='all').reset_index(drop=True)

    df = df.iloc[start_row[0]:end_row[0] - 3, :].reset_index(drop=True)
    df = df.dropna(axis=0, how='all')
    df = df.dropna(axis=1, how='all')
    df.columns = df.iloc[0]
    df = df.drop(0)
    df = df.drop(1)
    return df, df_head


def groupby_df(df):
    df = df.reset_index(drop=True)
    if 'Замена' in df.columns:
        df = df.groupby(['Номер', 'Описание', 'Вес детали'], as_index=False).agg({
            'Замена': 'max',
            'Марка': 'max',
            'Reference': 'max',
            'Кол-во': 'sum',
            'Цена RUB': 'mean',
            'Сумма RUB': 'sum',
            'Общий вес': 'sum',
        })
    else:
        df = df.groupby(['Номер', 'Описание', 'Вес детали'], as_index=False).agg({
            'Марка': 'max',
            'Reference': 'max',
            'Кол-во': 'sum',
            'Цена RUB': 'mean',
            'Сумма RUB': 'sum',
            'Общий вес': 'sum',
        })

    return df


def sort_df(df):
    if 'Замена' in df.columns:
        df = df[['Замена', 'Марка', 'Номер', 'Reference', 'Описание',
                 'Кол-во', 'Цена RUB', 'Сумма RUB', 'Вес детали', 'Общий вес']]
    else:
        df = df[['Марка', 'Номер', 'Reference', 'Описание',
                 'Кол-во', 'Цена RUB', 'Сумма RUB', 'Вес детали', 'Общий вес']]
    return df


def final_scores(df):
    total_sum = round(df['Сумма RUB'].sum(), 2)
    total_weight = round(df['Общий вес'].sum(), 3)
    cost_delivery = total_weight * 200
    return total_sum, total_weight, cost_delivery


def df_to_excel(df, header, file_name):

    start_row_table = 4
    # Получаем итоговые значения
    logger.info('Считаем итоговые значения')
    total_sum, total_weight, cost_delivery = final_scores(df)

    # Сбрасываем встроенный формат заголовков pandas
    pd.io.formats.excel.ExcelFormatter.header_style = None

    # Открываем файл для записи
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    sheet_name = 'Sheet1'  # Задаём имя вкладки
    workbook = writer.book  # Открываем книгу для записи

    # Записываем данные на вкладку sheet_name
    header.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=1)
    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row_table)

    # Выбираем вкладку для форматирования
    wks1 = writer.sheets[sheet_name]

    header_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '7',
        'align': 'center',
        'valign': 'top',
        'text_wrap': True,
        'bold': True,
        'bg_color': '#F4ECC5',
        'border': True,
        'border_color': '#CCC085'
    })
    # Форматируем таблицу
    wks1.set_default_row(12)
    wks1.set_row(start_row_table, 20, header_format)
    # wks1.set_column('A:A', 12, name_format)
    # wks1.set_column('B:B', 32, name_format)
    # wks1.set_column('C:H', 10, data_format)
    # wks1.set_column('I:I', 12, data_format)
    name_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'left',
        'valign': 'top',
        'text_wrap': True,
        'bold': False,
        'border': True,
        'border_color': '#CCC085'
    })
    name_format_rigth = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'rigth',
        'valign': 'top',
        'text_wrap': True,
        'bold': False,
        'border': True,
        'border_color': '#CCC085'
    })
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 10,
        'font_name': 'Arial',
        'valign': 'left'
    })
    footer_format = workbook.add_format({
        'bold': True,
        'font_size': 10,
        'font_name': 'Arial',
        'valign': 'right'
    })
    wks1.set_column('A:F', 10, name_format)
    wks1.set_column('F4:J100', 10, name_format_rigth)

    end_table = start_row_table + len(df) + 1

    # Записываем итоговые данные
    wks1.write(f'F{end_table + 2}', 'Итого:', None)
    wks1.write(f'G{end_table + 2}', total_sum, None)
    wks1.write(f'F{end_table + 4}', 'Итого за вес:', None)
    wks1.write(f'G{end_table + 4}', cost_delivery, None)
    wks1.write(f'H{end_table + 2}', 'Итого вес:', None)
    wks1.write(f'I{end_table + 2}', total_weight, None)

    # Изменяем формат строк заголовка
    for i in range(0, start_row_table):
        wks1.set_row(i, 16, header_format)
    # Изменяем формат строк подвала таблицы
    for i in range(end_table, end_table + 200):
        wks1.set_row(i, 16, footer_format)

# Сохраняем файл
    writer.save()
    return



def run():
    logger.info('Получаем имя файла')
    filename = search_file()
    logger.info('Считываем файл в память')
    df = read_file(filename)
    logger.info('Перестраиваем таблицу')
    df, df_header = rebuild_df(df)
    logger.info('Группируем значения по Артикулу')
    df = groupby_df(df)
    logger.info('Подготавливаем таблицу для записи')
    df = sort_df(df)
    logger.info('Сохраняем в файл')
    new_filename = f'_{filename[:-4]}_.xlsx'
    df_to_excel(df, df_header, new_filename)
    logger.info('Программа завершила свою работу')


if __name__ == '__main__':
    run()
