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
    mask_start = df == 'Country of'
    start_row = df[mask_start].dropna(axis=0, how='all').index.values

    df = df.iloc[start_row[0]:end_row[0] - 3, :].reset_index(drop=True)
    df = df.dropna(axis=0, how='all')
    df = df.dropna(axis=1, how='all')
    df.columns = df.iloc[0]
    df = df.drop(0)
    df = df.drop(1)
    return df

def groupby_df(df):
    df = df.reset_index(drop=True)
    df = df[['Country of', 'Марка', 'Номер', 'Описание', 'Кол-во', 'Цена $', 'Цена RUB', 'Сумма RUB', 'Вес детали']]
    df = df.groupby(['Номер', 'Описание', 'Цена $', 'Цена RUB', 'Вес детали']).sum()
    return df



def run():
    logger.info('Получаем имя файла')
    filename = search_file()
    logger.info('Считываем файл в память')
    df = read_file(filename)
    logger.info('Перестраиваем таблицу')
    df = rebuild_df(df)
    logger.info('Группируем значения по Артикулу')
    df = groupby_df(df)
    logger.info('Сохраняем в файл')
    df.to_excel('Сгруппированный заказ.xlsx')


if __name__ == '__main__':
    run()