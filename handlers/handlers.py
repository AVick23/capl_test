from telegram import Update
from telegram.ext import (
    ContextTypes,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters
)
from .constants import *
from .keyboards import *
from .utils import *
import logging

logger = logging.getLogger(__name__)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик команды /start"""
    # Если это сообщение (не callback)
    if update.message:
        await update.message.reply_text(
            START_MESSAGE,
            reply_markup=main_menu_keyboard()
        )
    # Если это callback (например, после /start в conversation)
    elif update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text(
            START_MESSAGE,
            reply_markup=main_menu_keyboard()
        )
    return MENU


async def main_menu_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Возврат в главное меню"""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        WELCOME_BACK,
        reply_markup=main_menu_keyboard()
    )
    return MENU


async def add_data_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начало добавления данных"""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        SELECT_METHOD,
        reply_markup=method_keyboard()
    )
    return METHOD


async def method_manual(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Ручной ввод - запрос даты"""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        ENTER_DATE,
        reply_markup=back_keyboard()
    )
    return DATE


async def method_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Загрузка фото"""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        "🖼 Отправьте фото отчёта",
        reply_markup=back_keyboard()
    )
    return PHOTO


async def process_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка даты"""
    date_str = update.message.text.strip()
    date_formatted, day_of_week = parse_date(date_str)
    
    if not date_formatted:
        await update.message.reply_text(
            f"{INVALID_FORMAT}\n\n{ENTER_DATE}",
            reply_markup=back_keyboard()
        )
        return DATE
    
    context.user_data['date'] = date_formatted
    context.user_data['day_of_week'] = day_of_week
    
    await update.message.reply_text(
        f"✅ Дата: {date_formatted} ({day_of_week})\n\n{ENTER_SALES}",
        reply_markup=back_keyboard()
    )
    return SALES


async def process_sales(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка продаж"""
    try:
        sales = float(update.message.text.replace(',', '.'))
        if sales < 0:
            raise ValueError
        
        context.user_data['sales'] = sales
        await update.message.reply_text(
            f"✅ Продажи: {sales:,.0f}\n\n{ENTER_CHECKS}",
            reply_markup=back_keyboard()
        )
        return CHECKS
    except ValueError:
        await update.message.reply_text(
            f"{INVALID_FORMAT}\n\n{ENTER_SALES}",
            reply_markup=back_keyboard()
        )
        return SALES


async def process_checks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка чеков"""
    try:
        checks = int(update.message.text)
        if checks < 0:
            raise ValueError
        
        context.user_data['checks'] = checks
        await update.message.reply_text(
            f"✅ Чеки: {checks}\n\n{ENTER_COST_PRICE}",
            reply_markup=back_keyboard()
        )
        return COST_PRICE
    except ValueError:
        await update.message.reply_text(
            f"{INVALID_FORMAT}\n\n{ENTER_CHECKS}",
            reply_markup=back_keyboard()
        )
        return CHECKS


async def process_cost_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка себестоимости"""
    try:
        cost_price = float(update.message.text.replace(',', '.'))
        if cost_price < 0:
            raise ValueError
        
        context.user_data['cost_price'] = cost_price
        await update.message.reply_text(
            f"✅ Себестоимость: {cost_price:,.0f}\n\n{ENTER_EXPENSES}",
            reply_markup=back_keyboard()
        )
        return EXPENSES
    except ValueError:
        await update.message.reply_text(
            f"{INVALID_FORMAT}\n\n{ENTER_COST_PRICE}",
            reply_markup=back_keyboard()
        )
        return COST_PRICE


async def process_expenses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка расходов"""
    try:
        expenses = float(update.message.text.replace(',', '.'))
        if expenses < 0:
            raise ValueError
        
        context.user_data['expenses'] = expenses
        
        # Показываем подтверждение
        await show_confirmation(update, context)
        return CONFIRM
    except ValueError:
        await update.message.reply_text(
            f"{INVALID_FORMAT}\n\n{ENTER_EXPENSES}",
            reply_markup=back_keyboard()
        )
        return EXPENSES


async def show_confirmation(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Показать подтверждение данных"""
    data = context.user_data
    
    text = f"""{CONFIRM_DATA}

📅 {DATE_LABEL}: {data.get('date', 'N/A')} ({data.get('day_of_week', '')})
💰 {SALES_LABEL}: {data.get('sales', 0):,.0f}
🧾 {CHECKS_LABEL}: {data.get('checks', 0)}
💸 {COST_PRICE_LABEL}: {data.get('cost_price', 0):,.0f}
📉 {EXPENSES_LABEL}: {data.get('expenses', 0):,.0f}
"""
    
    await update.message.reply_text(
        text,
        reply_markup=confirm_keyboard()
    )


async def process_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка фото"""
    await update.message.reply_text(PROCESSING_PHOTO)
    
    # Получаем фото
    photo = update.message.photo[-1]
    file = await photo.get_file()
    photo_bytes = await file.download_as_bytearray()
    
    # OCR
    values = extract_from_photo(photo_bytes)
    
    if values and len(values) >= 4:
        # Попытка распознать данные
        try:
            context.user_data['sales'] = float(values[0])
            context.user_data['checks'] = int(values[1]) if len(values) > 1 else 0
            context.user_data['cost_price'] = float(values[2]) if len(values) > 2 else 0
            context.user_data['expenses'] = float(values[3]) if len(values) > 3 else 0
            
            await update.message.reply_text(
                f"✅ Данные распознаны!\n\n{ENTER_DATE}",
                reply_markup=back_keyboard()
            )
            return DATE
        except (ValueError, IndexError) as e:
            logger.error(f"Error parsing photo values: {e}")
            await update.message.reply_text(
                f"{PHOTO_FAILED}\n\n{ENTER_DATE}",
                reply_markup=back_keyboard()
            )
            return DATE
    else:
        await update.message.reply_text(
            f"{PHOTO_FAILED}\n\n{ENTER_DATE}",
            reply_markup=back_keyboard()
        )
        return DATE


async def confirm_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Подтверждение и сохранение данных"""
    query = update.callback_query
    await query.answer()
    
    data = context.user_data
    
    # Сохраняем в БД
    success = add_data_to_db(
        data['date'],
        data['day_of_week'],
        data['sales'],
        data['checks'],
        data['cost_price'],
        data['expenses']
    )
    
    if success:
        await query.edit_message_text(
            f"{DATA_SAVED}\n\nЧто хотите сделать дальше?",
            reply_markup=main_menu_keyboard()
        )
    else:
        await query.edit_message_text(
            "❌ Ошибка сохранения. Попробуйте ещё раз.",
            reply_markup=main_menu_keyboard()
        )
    
    # Очищаем данные
    context.user_data.clear()
    
    return MENU


async def edit_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Редактирование данных"""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        EDIT_PROMPT,
        reply_markup=edit_keyboard()
    )
    return EDIT


async def edit_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Редактирование конкретного поля"""
    query = update.callback_query
    await query.answer()
    
    field = query.data.replace('edit_', '')
    context.user_data['editing_field'] = field
    
    messages = {
        'date': ENTER_DATE,
        'sales': ENTER_SALES,
        'checks': ENTER_CHECKS,
        'cost_price': ENTER_COST_PRICE,
        'expenses': ENTER_EXPENSES
    }
    
    states = {
        'date': DATE,
        'sales': SALES,
        'checks': CHECKS,
        'cost_price': COST_PRICE,
        'expenses': EXPENSES
    }
    
    await query.edit_message_text(
        messages[field],
        reply_markup=back_keyboard()
    )
    
    return states[field]


async def back_to_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Возврат к подтверждению"""
    query = update.callback_query
    await query.answer()
    
    # Создаём фиктивное сообщение для show_confirmation
    await query.edit_message_text(
        f"""{CONFIRM_DATA}

📅 {DATE_LABEL}: {context.user_data.get('date', 'N/A')} ({context.user_data.get('day_of_week', '')})
💰 {SALES_LABEL}: {context.user_data.get('sales', 0):,.0f}
🧾 {CHECKS_LABEL}: {context.user_data.get('checks', 0)}
💸 {COST_PRICE_LABEL}: {context.user_data.get('cost_price', 0):,.0f}
📉 {EXPENSES_LABEL}: {context.user_data.get('expenses', 0):,.0f}
""",
        reply_markup=confirm_keyboard()
    )
    return CONFIRM


async def view_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Просмотр отчёта"""
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text("⏳ Генерирую отчёт...")
    
    filename = generate_excel_report()
    
    if filename:
        with open(filename, 'rb') as f:
            await query.message.reply_document(
                document=f,
                caption=REPORT_GENERATED
            )
        await query.edit_message_text(
            "✅ Отчёт отправлен выше ☝️",
            reply_markup=main_menu_keyboard()
        )
    else:
        await query.edit_message_text(
            NO_DATA,
            reply_markup=main_menu_keyboard()
        )
    
    return MENU


async def view_analytics(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Просмотр аналитики"""
    query = update.callback_query
    await query.answer()
    
    await query.edit_message_text("⏳ Генерирую аналитику...")
    
    filename, analytics = generate_dashboard()
    
    if filename:
        with open(filename, 'rb') as f:
            await query.message.reply_photo(
                photo=f,
                caption=analytics
            )
        await query.edit_message_text(
            "✅ Аналитика отправлена выше ☝️",
            reply_markup=main_menu_keyboard()
        )
    else:
        await query.edit_message_text(
            NO_DATA,
            reply_markup=main_menu_keyboard()
        )
    
    return MENU


async def set_targets(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Установка целей"""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        TARGETS_INSTRUCTION,
        reply_markup=back_keyboard()
    )
    return TARGETS


async def process_targets(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка целей"""
    try:
        text = update.message.text.strip()
        values = text.split()
        
        if len(values) != 3:
            raise ValueError
        
        sales_target = float(values[0])
        net_profit_target = float(values[1])
        expenses_target = float(values[2])
        
        # Получаем текущий месяц и год
        now = datetime.now()
        month = now.strftime('%B')
        year = now.year
        
        set_targets_db(sales_target, net_profit_target, expenses_target, month, year)
        
        await update.message.reply_text(
            f"{TARGETS_SET}\n\nЧто хотите сделать дальше?",
            reply_markup=main_menu_keyboard()
        )
        return MENU
    except ValueError:
        await update.message.reply_text(
            f"{INVALID_FORMAT}\n\n{TARGETS_INSTRUCTION}",
            reply_markup=back_keyboard()
        )
        return TARGETS


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отмена действия"""
    context.user_data.clear()
    
    if update.callback_query:
        query = update.callback_query
        await query.answer()
        await query.edit_message_text(
            CANCELLED,
            reply_markup=main_menu_keyboard()
        )
    else:
        await update.message.reply_text(
            CANCELLED,
            reply_markup=main_menu_keyboard()
        )
    
    return MENU


async def back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Кнопка назад"""
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        SELECT_METHOD,
        reply_markup=method_keyboard()
    )
    return METHOD


async def back_to_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Возврат в главное меню из любого места"""
    query = update.callback_query
    await query.answer()
    context.user_data.clear()
    await query.edit_message_text(
        MAIN_MENU_TEXT,
        reply_markup=main_menu_keyboard()
    )
    return MENU


async def unknown_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработчик неизвестных callback'ов"""
    query = update.callback_query
    await query.answer("⚠️ Неизвестная команда")
    logger.warning(f"Unknown callback: {query.data}")
    await query.edit_message_text(
        "⚠️ Команда не распознана. Пожалуйста, нажмите /start для перезапуска.",
        reply_markup=main_menu_keyboard()
    )
    return MENU


def setup_handlers(application):
    """Регистрация всех обработчиков"""
    
    # ConversationHandler для добавления данных
    conv_handler = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(add_data_entry, pattern='^add_data$')
        ],
        states={
            MENU: [
                # 🔥 КЛЮЧЕВОЕ ИСПРАВЛЕНИЕ
                CallbackQueryHandler(add_data_entry, pattern='^add_data$'),
                CallbackQueryHandler(main_menu_callback, pattern='^main_menu$'),
                CallbackQueryHandler(view_report, pattern='^view_report$'),
                CallbackQueryHandler(view_analytics, pattern='^analytics$'),
                CallbackQueryHandler(set_targets, pattern='^set_targets$')
            ],
            METHOD: [
                CallbackQueryHandler(method_manual, pattern='^method_manual$'),
                CallbackQueryHandler(method_photo, pattern='^method_photo$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
                # На случай повторного нажатия кнопок меню
                CallbackQueryHandler(add_data_entry, pattern='^add_data$'),
                CallbackQueryHandler(view_report, pattern='^view_report$'),
                CallbackQueryHandler(view_analytics, pattern='^analytics$'),
                CallbackQueryHandler(set_targets, pattern='^set_targets$')
            ],
            DATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_date),
                CallbackQueryHandler(back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$')
            ],
            SALES: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_sales),
                CallbackQueryHandler(back, pattern='^back$')
            ],
            CHECKS: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_checks),
                CallbackQueryHandler(back, pattern='^back$')
            ],
            COST_PRICE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_cost_price),
                CallbackQueryHandler(back, pattern='^back$')
            ],
            EXPENSES: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_expenses),
                CallbackQueryHandler(back, pattern='^back$')
            ],
            PHOTO: [
                MessageHandler(filters.PHOTO, process_photo),
                CallbackQueryHandler(back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$')
            ],
            CONFIRM: [
                CallbackQueryHandler(confirm_data, pattern='^confirm$'),
                CallbackQueryHandler(edit_data, pattern='^edit$'),
                CallbackQueryHandler(cancel, pattern='^cancel_add$')
            ],
            EDIT: [
                CallbackQueryHandler(edit_field, pattern='^edit_'),
                CallbackQueryHandler(back_to_confirm, pattern='^back_to_confirm$')
            ],
            TARGETS: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_targets),
                CallbackQueryHandler(back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$')
            ]
        },
        fallbacks=[
            CommandHandler('cancel', cancel),
            CommandHandler('start', start)
        ]
    )
    
    application.add_handler(CommandHandler('start', start))
    application.add_handler(conv_handler)
    
    # Fallback для неизвестных callbacks
    application.add_handler(CallbackQueryHandler(unknown_callback))