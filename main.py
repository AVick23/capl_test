import logging
from telegram.ext import ApplicationBuilder, CommandHandler
from config import BOT_TOKEN
from handlers import setup_handlers
from handlers.utils import init_db

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)

logger = logging.getLogger(__name__)


async def error_handler(update, context):
    """Обработчик ошибок"""
    logger.error(f"Exception: {context.error}")
    if update:
        await update.effective_message.reply_text(
            "❌ Произошла ошибка. Пожалуйста, попробуйте ещё раз."
        )


def main():
    """Запуск бота"""
    # Инициализация БД
    init_db()
    
    # Создание приложения
    application = ApplicationBuilder().token(BOT_TOKEN).build()
    
    # Регистрация обработчиков ошибок
    application.add_error_handler(error_handler)
    
    # Регистрация всех обработчиков
    setup_handlers(application)
    
    # Запуск
    logger.info("Бот запущен")
    application.run_polling()


if __name__ == "__main__":
    main()