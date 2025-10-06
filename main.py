import os
import shutil
import asyncio
import logging
from aiogram import Bot, Dispatcher, types, F, Router
from aiogram.filters import CommandStart
from aiogram.types import FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.storage.memory import MemoryStorage
from dotenv import load_dotenv

# Импортируем функции из наших новых файлов
from file_processing import process_excel_files
from email_sender import send_email_with_attachment

# --- НАСТРОЙКА ЛОГИРОВАНИЯ ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# --- ЗАГРУЗКА ПЕРЕМЕННЫХ ИЗ .ENV ---
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")
EMAIL_TO = os.getenv("EMAIL_TO")
TEMP_FOLDER = "temp_files"

# --- ИНИЦИАЛИЗАЦИЯ БОТА ---
storage = MemoryStorage()
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=storage)
router = Router()


# --- ОБРАБОТЧИКИ СООБЩЕНИЙ ---
@router.message(CommandStart())
async def send_welcome(message: types.Message):
    await message.answer(
        "Здравствуйте! Отправьте мне файл с остатками из МойСклад."
    )


@router.message(F.document)
async def handle_document(message: types.Message):
    if not message.document.file_name.endswith(('.xls', '.xlsx')):
        logging.warning("Получен файл неверного формата.")
        return await message.answer("Пожалуйста, отправьте файл в формате .xls или .xlsx.")

    logging.info(f"Получен файл от пользователя {message.from_user.id}: {message.document.file_name}")
    await message.answer("Файл получен, начинаю обработку...")

    if not os.path.exists(TEMP_FOLDER):
        os.makedirs(TEMP_FOLDER)
        logging.info(f"Создана временная папка: {TEMP_FOLDER}")

    output_file_name = None
    try:
        file_path = os.path.join(TEMP_FOLDER, message.document.file_name)
        await bot.download(message.document, destination=file_path)
        logging.info(f"Файл сохранен во временную папку: {file_path}")

        output_file_name = process_excel_files(file_path)

        if os.path.exists(output_file_name):
            await message.answer_document(FSInputFile(output_file_name), caption="Вот ваш обработанный файл.")

            keyboard = InlineKeyboardMarkup(inline_keyboard=[
                [
                    InlineKeyboardButton(text="✅", callback_data="send_email_yes"),
                    InlineKeyboardButton(text="❌", callback_data="send_email_no")
                ]
            ])
            await message.answer(
                f"Отправить на почту {EMAIL_TO}?",
                reply_markup=keyboard
            )
        else:
            await message.answer(f"Произошла ошибка при обработке файла: {output_file_name}")

    except Exception as e:
        logging.error(f"Критическая ошибка в handle_document: {e}", exc_info=True)
        await message.answer(f"Произошла ошибка: {e}")
    finally:
        pass


@router.callback_query(F.data.startswith('send_email_'))
async def handle_email_request(callback_query: types.CallbackQuery):
    output_file_name = None
    try:
        message = callback_query.message
        user_choice = callback_query.data.split('_')[-1]

        await callback_query.answer()

        file_list = [f for f in os.listdir('.') if
                     f.startswith('Остатки ИП Лесковский') and f != 'Остатки ИП Лесковский.xlsx']
        if file_list:
            output_file_name = max(file_list, key=os.path.getmtime)

        if user_choice == 'yes':
            if output_file_name and os.path.exists(output_file_name):
                if await send_email_with_attachment(output_file_name):
                    await message.answer(f"Файл успешно отправлен на почту {EMAIL_TO}.")
                else:
                    await message.answer("Не удалось отправить файл. Проверьте логи.")
            else:
                logging.warning(f"Файл для отправки не найден: {output_file_name}")
                await message.answer("Ошибка: не удалось найти файл для отправки.")
        else:
            logging.info("Пользователь выбрал 'нет'. Отправка отменена.")
            await message.answer("Отмена отправки. Файл не будет отправлен на почту.")

    except Exception as e:
        logging.error(f"Критическая ошибка в handle_email_request: {e}", exc_info=True)
        await message.answer(f"Произошла ошибка при обработке запроса: {e}")
    finally:
        if os.path.exists(TEMP_FOLDER):
            shutil.rmtree(TEMP_FOLDER)
            logging.info(f"Временная папка '{TEMP_FOLDER}' удалена.")
        if output_file_name and os.path.exists(output_file_name):
            os.remove(output_file_name)
            logging.info(f"Обработанный файл '{output_file_name}' удален.")


# --- ЗАПУСК БОТА ---
async def main() -> None:
    dp.include_router(router)
    if BOT_TOKEN is None:
        logging.error("Ошибка: Токен бота не найден. Убедитесь, что BOT_TOKEN есть в файле .env")
        print(
            "Ошибка: Токен бота не найден. Пожалуйста, убедитесь, что вы создали файл .env и добавили в него BOT_TOKEN.")
        return
    logging.info("Бот запущен. Ожидание сообщений...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
    