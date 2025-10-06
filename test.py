import requests

# Данные для доступа к API RetailCRM
RETAILCRM_BASE_URL = "https://tropichouse.retailcrm.ru"
RETAILCRM_API_KEY = "Ml4gDpj0AycwsR6JmTM7wisCX3kA5AlB"
RETAILCRM_SITE_CODE = "tropichouse"


def get_order_info(order_id: str):
    """
    Получает информацию о заказе по его внутреннему ID из RetailCRM.

    Args:
        order_id (str): Внутренний ID заказа.
    """
    api_url = f"{RETAILCRM_BASE_URL}/api/v5/orders/{order_id}"

    # Параметры запроса, включая API-ключ, код магазина и параметр "by=id"
    params = {
        'apiKey': RETAILCRM_API_KEY,
        'site': RETAILCRM_SITE_CODE,
        'by': 'id'  # ⬅️ Указываем, что передаем внутренний ID
    }

    try:
        # Отправляем GET-запрос к API
        response = requests.get(api_url, params=params)

        # Проверяем, что запрос был успешен
        response.raise_for_status()

        # Парсим ответ в формате JSON
        data = response.json()

        # Проверяем, что API вернул успешный результат
        if data.get('success'):
            order_data = data.get('order')
            if order_data:
                print("Информация о заказе успешно получена:")
                # Выводим информацию о заказе в читаемом формате
                for key, value in order_data.items():
                    print(f"- {key}: {value}")
                print(order_data)
            else:
                print(f"Заказ с ID {order_id} не найден.")
        else:
            print("Ошибка API:", data.get('errorMsg'))

    except requests.exceptions.RequestException as e:
        print(f"Ошибка при выполнении запроса: {e}")
    except ValueError as e:
        print(f"Ошибка при разборе JSON-ответа: {e}")


# Пример использования:
if __name__ == "__main__":
    order_id = "25077"  # Внутренний ID заказа
    get_order_info(order_id)