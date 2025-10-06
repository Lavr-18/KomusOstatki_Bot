import os
import requests
import re
import logging
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()
CRM_API_URL = os.getenv("RETAILCRM_BASE_URL")
CRM_API_KEY = os.getenv("RETAILCRM_API_KEY")

CRM_API_VERSION = "v5"


def get_crm_popularity():
    """
    Получает список заказов из RetailCRM за последние 2 месяца со способом 'komus'
    и рассчитывает популярность каждой композиции.

    Возвращает: dict, где ключ - название композиции, значение - количество заказов.
    """
    if not CRM_API_URL or not CRM_API_KEY:
        logging.error("Ошибка: Не найдены переменные окружения RETAILCRM_BASE_URL или RETAILCRM_API_KEY.")
        return {}

    logging.info("Начало запроса популярности из RetailCRM.")

    # Расчет даты 2 месяца назад для фильтрации
    date_from = (datetime.now() - timedelta(days=60)).strftime('%Y-%m-%d')

    # Базовый URL для API
    base_url = f"{CRM_API_URL}/api/{CRM_API_VERSION}/orders"

    # Параметры фильтра
    params = {
        'apiKey': CRM_API_KEY,
        'filter[orderMethods][]': 'komus',  # Способ оформления Комус
        'filter[createdAtFrom]': date_from,  # За последние 2 месяца
        'limit': 100,  # Максимальный лимит для постранички
        'page': 1
    }

    all_orders = []
    total_pages = 1
    set_name_pattern = re.compile(r'\[(.*?)\]')  # Регулярное выражение для извлечения [Текст]

    try:
        # Цикл для обхода постраничной разбивки
        while params['page'] <= total_pages:
            response = requests.get(base_url, params=params)
            response.raise_for_status()  # Вызывает исключение для HTTP ошибок
            data = response.json()

            if not data.get('success'):
                logging.error(f"Ошибка в ответе CRM: {data.get('errorMsg', 'Неизвестная ошибка')}")
                return {}

            if 'pagination' in data:
                total_pages = data['pagination']['totalPageCount']

            all_orders.extend(data.get('orders', []))

            logging.info(
                f"Загружена страница {params['page']} из {total_pages}. Заказов: {len(data.get('orders', []))}")
            params['page'] += 1

    except requests.exceptions.RequestException as err:
        logging.error(f"Ошибка при запросе к CRM: {err}")
        return {}
    except Exception as e:
        logging.error(f"Непредвиденная ошибка при работе с CRM: {e}", exc_info=True)
        return {}

    logging.info(f"Всего загружено {len(all_orders)} заказов. Начинаю расчет популярности.")

    popularity_map = {}

    for order in all_orders:
        unique_sets_in_order = set()

        for item in order.get('items', []):
            properties = item.get('properties')
            set_name_value = ''

            # --- ИСПРАВЛЕННАЯ ЛОГИКА ДЛЯ ОБРАБОТКИ СПИСКА ИЛИ СЛОВАРЯ ---
            if isinstance(properties, list):
                # Если properties - это список объектов (как предполагается из ошибки)
                for prop in properties:
                    if prop.get('code') == 'SET_NAME':
                        set_name_value = prop.get('value', '')
                        break
            elif isinstance(properties, dict):
                # Если properties - это словарь, где ключи - коды свойств (изначальная логика)
                set_name_value = properties.get('SET_NAME', {}).get('value', '')
            # --- КОНЕЦ ИСПРАВЛЕННОЙ ЛОГИКИ ---

            if set_name_value:
                # Извлекаем текст внутри квадратных скобок [...]
                match = set_name_pattern.search(set_name_value)
                if match:
                    composition_name = match.group(1).strip()
                    # Добавляем композицию в набор (уникальность на заказ)
                    unique_sets_in_order.add(composition_name)

        # Обновляем счетчик популярности
        for name in unique_sets_in_order:
            popularity_map[name] = popularity_map.get(name, 0) + 1

    logging.info(f"Расчет популярности завершен. Найдено {len(popularity_map)} уникальных композиций.")
    return popularity_map