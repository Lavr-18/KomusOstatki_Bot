import os
import shutil
import pandas as pd
import logging
import numpy as np
import math
from datetime import datetime

# --- ИМПОРТ ИЗ НОВОГО МОДУЛЯ ---
from crm_connector import get_crm_popularity

# --- КОНСТАНТЫ ---
FILE_OST_KOMUS = 'Остатки_Комус.xlsx'
FILE_OST_LESKOVSKY = 'Остатки ИП Лесковский.xlsx'
TEMP_FOLDER = "temp_files"

REPORT_RANGE = 'D13:F83'
REPORT_COPY_RANGE = 'F13:F74'
KOMUS_LIST_1 = 'счет остатков new (копия)'
KOMUS_COLUMNS_1 = [0, 1, 2]
KOMUS_LIST_2 = 'Состав композиций (копия)'
# Загружаем колонки: [2] Состав композиции (для calculate_formulas), [4] Название в СРМ, [5] Название растений
KOMUS_COLUMNS_FOR_ROUNDING = [2, 4, 5]
COL_INDEX_COMPOSITION_KEY = 0  # Индекс 0 в новом фрейме komus_df_2_all (Состав композиции)
COL_INDEX_CRM_NAME = 1  # Индекс 1 в новом фрейме komus_df_2_all (Название в СРМ)
COL_INDEX_PLANT_NAME = 2  # Индекс 2 в новом фрейме komus_df_2_all (Название растений)
LESKOVSKY_SHEET = 'Лист1'


def calculate_formulas(komus_sheet1_df, komus_sheet2_df):
    """
    Вычисляет значения для столбца D, имитируя формулы Excel.
    Использует колонку "Состав композиции", которая теперь имеет индекс 0 в komus_sheet2_df.
    """
    results = []
    # komus_sheet1_df.iloc[:, 1] - Артикул, komus_sheet1_df.iloc[:, 2] - Остаток (VPR)
    vlookup_dict = dict(zip(komus_sheet1_df.iloc[:, 1], komus_sheet1_df.iloc[:, 2]))

    # Колонка с ключом - это индекс 0 в komus_sheet2_df, т.е. исходный столбец 2
    countif_dict = komus_sheet2_df.iloc[:, COL_INDEX_COMPOSITION_KEY].value_counts().to_dict()

    logging.info("Начало расчёта формул.")

    for i, item in enumerate(komus_sheet2_df.iloc[:, COL_INDEX_COMPOSITION_KEY]):
        vlookup_result = vlookup_dict.get(item)
        countif_result = countif_dict.get(item)

        # Обработка NaN/None в VPR
        if vlookup_result is None or pd.isna(vlookup_result):
            vlookup_result = 0

        logging.info(f"Строка {i + 2}: Артикул: '{item}', ВПР: {vlookup_result}, СЧЁТЕСЛИ: {countif_result}")

        if countif_result is None or countif_result == 0:
            results.append(0)
            logging.info(f"   --> Результат: 0 (ошибка или нулевое значение)")
        else:
            result = vlookup_result / countif_result
            results.append(result)
            logging.info(f"   --> Результат: {result}")

    return results


def apply_rounding_logic(df_data, calculated_data, popularity_map):
    """
    Применяет сложную логику округления: группировка по растению,
    суммирование дробных частей, распределение целого бонуса по популярности.
    """
    logging.info("Начало применения сложной логики округления.")

    # 1. Создаем рабочий DataFrame
    df = df_data.copy()
    df['Calculated_Value'] = calculated_data
    df['Final_Value'] = 0.0

    # Обнуляем отрицательные значения и фильтруем, оставляя только те, где остаток > 0
    df['Calculated_Value'] = df['Calculated_Value'].apply(lambda x: max(0, x))
    df_positive = df[df['Calculated_Value'] > 0].copy()

    if df_positive.empty:
        # Возвращаем серию нулей, сохраняя длину и порядок исходных данных
        return pd.Series(np.zeros(len(calculated_data)), dtype=float)

        # 2. Добавляем популярность (0, если нет в CRM)
    df_positive['Popularity'] = df_positive.iloc[:, COL_INDEX_CRM_NAME].apply(lambda x: popularity_map.get(x, 0))

    # 3. Группируем по Названию растений
    grouped = df_positive.groupby(df_positive.columns[COL_INDEX_PLANT_NAME])

    # 4. Обработка каждой группы
    for plant_name, group in grouped:

        # a. Расчет бонуса
        # Получаем дробные части: X.YY -> 0.YY
        fractional_parts = group['Calculated_Value'].apply(lambda x: x - math.floor(x))
        total_fraction = fractional_parts.sum()
        # Округляем сумму вниз для получения целого бонуса
        integer_bonus = math.floor(total_fraction)

        logging.info(
            f"Группа '{plant_name}': Композиций: {len(group)}, Сумма дробных частей: {total_fraction:.2f}, Бонус: {integer_bonus}")

        # b. Сортировка по популярности (убывание)
        # Внутри группы сортируем по популярности, при равной популярности - по исходному значению (для тай-брейка)
        group_sorted = group.sort_values(by=['Popularity', 'Calculated_Value'], ascending=[False, False]).copy()

        # c. Распределение бонуса
        for index, row in group_sorted.iterrows():
            floor_value = math.floor(row['Calculated_Value'])
            bonus = 0

            if integer_bonus > 0:
                # Если бонус есть, даем +1 самой популярной/крупной композиции
                bonus = 1
                integer_bonus -= 1

            # d. Определение финального значения (целая часть + бонус)
            group_sorted.loc[index, 'Final_Value'] = floor_value + bonus

        # e. Обновление основного DataFrame с округленными значениями
        df_positive.loc[group_sorted.index, 'Final_Value'] = group_sorted['Final_Value']

    # 5. Форматирование результата в серию, соответствующую длине и порядку исходных данных
    # Создаем серию нулей, соответствующую длине и порядку исходных данных (df_data)
    result_series = pd.Series(np.zeros(len(calculated_data)), dtype=float)

    # Обновляем значениями только те индексы, которые были обработаны (где > 0)
    for original_index in df_positive.index:
        result_series.loc[original_index] = df_positive.loc[original_index, 'Final_Value']

    logging.info("Округление завершено.")
    return result_series


def process_excel_files(input_file_path):
    """
    Основная логика обработки файлов, теперь включающая получение данных из CRM
    и сложную логику округления.
    """
    logging.info(f"Начало обработки файла: {input_file_path}")
    try:
        # 1. ЗАГРУЗКА ДАННЫХ ИЗ ФАЙЛОВ
        start_col = REPORT_RANGE.split(':')[0][0]
        end_col = REPORT_RANGE.split(':')[1][0]
        start_row = int(''.join(filter(str.isdigit, REPORT_RANGE.split(':')[0])))
        end_row = int(''.join(filter(str.isdigit, REPORT_RANGE.split(':')[1])))

        # Чтение данных из входного файла (для VPR)
        report_df = pd.read_excel(
            input_file_path,
            header=None,
            usecols=f'{start_col}:{end_col}',
            skiprows=start_row - 1,
            nrows=end_row - start_row
        )
        report_df_sorted = report_df.sort_values(by=report_df.columns[0]).reset_index(drop=True)
        start_row_f = int(''.join(filter(str.isdigit, REPORT_COPY_RANGE.split(':')[0]))) - 1 - (start_row - 1)
        end_row_f = int(''.join(filter(str.isdigit, REPORT_COPY_RANGE.split(':')[1]))) - (start_row - 1)
        data_from_report = report_df_sorted.iloc[start_row_f:end_row_f, 2]

        # Чтение данных из Остатки_Комус.xlsx
        komus_df_1 = pd.read_excel(FILE_OST_KOMUS, sheet_name=KOMUS_LIST_1, header=None, usecols=KOMUS_COLUMNS_1,
                                   skiprows=1)
        # Загружаем колонки [2, 4, 5] для calculate_formulas и apply_rounding_logic
        komus_df_2_all = pd.read_excel(FILE_OST_KOMUS, sheet_name=KOMUS_LIST_2, header=None,
                                       usecols=KOMUS_COLUMNS_FOR_ROUNDING,
                                       skiprows=1)

        # 2. ПОДГОТОВКА И РАСЧЕТ
        komus_df_1.iloc[:len(data_from_report), 2] = data_from_report.values
        calculated_data = calculate_formulas(komus_df_1, komus_df_2_all)

        # 3. НОВАЯ ЛОГИКА ОКРУГЛЕНИЯ
        # Получаем данные о популярности из CRM
        popularity_map = get_crm_popularity()

        # --- ДЕБАГ: Выводим словарь популярности ---
        logging.info(f"Словарь популярности (Название композиции: Кол-во заказов): {popularity_map}")

        # Применяем новую логику округления
        calculated_series = apply_rounding_logic(komus_df_2_all, calculated_data, popularity_map)

        # 4. ЗАПИСЬ В ФАЙЛ
        with pd.ExcelWriter(FILE_OST_LESKOVSKY, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            calculated_series.to_excel(
                writer,
                sheet_name=LESKOVSKY_SHEET,
                index=False,
                header=False,
                startrow=1,
                startcol=2
            )

        today_date = datetime.now().strftime('%d.%m')
        new_file_name = f'Остатки ИП Лесковский {today_date}.xlsx'
        shutil.copyfile(FILE_OST_LESKOVSKY, new_file_name)

        logging.info(f"Файл успешно обработан. Создан: {new_file_name}")
        return new_file_name

    except FileNotFoundError as e:
        logging.error(f"Ошибка: Файл не найден: {e}")
        return f"Ошибка: Файл не найден: {e}"
    except Exception as e:
        logging.error(f"Произошла ошибка при обработке: {e}", exc_info=True)
        return f"Произошла ошибка при обработке: {e}"