from telegram import Update
from telegram.ext import (
    ContextTypes,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters
)
from telegram.error import BadRequest
from .constants import *
from .keyboards import *
from .utils import *
import logging

logger = logging.getLogger(__name__)


# ============================================
# HELPER: Безопасное редактирование сообщения
# ============================================
async def safe_edit_message(query, text, reply_markup=None):
    """
    Безопасно редактирует сообщение, игнорируя ошибку
    'Message is not modified' (когда пытаемся изменить сообщение
    на идентичное).
    """
    try:
        await query.edit_message_text(text, reply_markup=reply_markup)
    except BadRequest as e:
        if "Message is not modified" not in str(e):
            raise  # Пробрасываем другие ошибки дальше
    except Exception as e:
        logger.error(f"Ошибка редактирования сообщения: {e}")


# ============================================
# ENTRY POINT: /start (также используется как fallback)
# ============================================
async def start_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Точка входа и сброс состояния при /start"""
    context.user_data.clear()
    
    # Если это callback_query (например, при нажатии кнопки), отвечаем
    if update.callback_query:
        query = update.callback_query
        await query.answer()
        await safe_edit_message(
            query,
            START_MESSAGE,
            reply_markup=main_menu_keyboard()
        )
    else:
        await update.message.reply_text(
            START_MESSAGE,
            reply_markup=main_menu_keyboard()
        )
    return MENU


# ============================================
# СОСТОЯНИЕ: MENU (главное меню)
# ============================================
async def menu_view_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отчёт из меню"""
    query = update.callback_query
    await query.answer()
    
    await safe_edit_message(query, "⏳ Генерирую отчёт...")
    
    filename = generate_excel_report()
    
    if filename:
        try:
            with open(filename, 'rb') as f:
                await query.message.reply_document(
                    document=f,
                    caption=REPORT_GENERATED
                )
            await safe_edit_message(
                query,
                "✅ Отчёт отправлен выше ☝️",
                reply_markup=main_menu_keyboard()
            )
        except Exception as e:
            logger.error(f"Ошибка отправки отчёта: {e}")
            await safe_edit_message(
                query,
                "❌ Ошибка при отправке. Попробуйте ещё раз.",
                reply_markup=main_menu_keyboard()
            )
    else:
        await safe_edit_message(
            query,
            NO_DATA,
            reply_markup=main_menu_keyboard()
        )
    
    return MENU


async def menu_view_analytics(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Аналитика из меню"""
    query = update.callback_query
    await query.answer()
    
    await safe_edit_message(query, "⏳ Генерирую аналитику...")
    
    filename, analytics = generate_dashboard()
    
    if filename:
        try:
            with open(filename, 'rb') as f:
                await query.message.reply_photo(
                    photo=f,
                    caption=analytics
                )
            await safe_edit_message(
                query,
                "✅ Аналитика отправлена выше ☝️",
                reply_markup=main_menu_keyboard()
            )
        except Exception as e:
            logger.error(f"Ошибка отправки аналитики: {e}")
            await safe_edit_message(
                query,
                "❌ Ошибка при отправке. Попробуйте ещё раз.",
                reply_markup=main_menu_keyboard()
            )
    else:
        await safe_edit_message(
            query,
            NO_DATA,
            reply_markup=main_menu_keyboard()
        )
    
    return MENU


async def menu_set_targets(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Установка целей из меню"""
    query = update.callback_query
    await query.answer()
    await safe_edit_message(
        query,
        TARGETS_INSTRUCTION,
        reply_markup=back_keyboard()
    )
    return TARGETS


async def menu_add_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Начало добавления данных из меню"""
    query = update.callback_query
    await query.answer()
    await safe_edit_message(
        query,
        SELECT_METHOD,
        reply_markup=method_keyboard()
    )
    return METHOD


# ============================================
# СОСТОЯНИЕ: METHOD (выбор метода ввода)
# ============================================
async def method_manual(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Ручной ввод"""
    query = update.callback_query
    await query.answer()
    await safe_edit_message(
        query,
        ENTER_DATE,
        reply_markup=back_keyboard()
    )
    return DATE


async def method_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Загрузка фото"""
    query = update.callback_query
    await query.answer()
    await safe_edit_message(
        query,
        "🖼 Отправьте фото отчёта",
        reply_markup=back_keyboard()
    )
    return PHOTO


# ============================================
# СОСТОЯНИЕ: DATE (ввод даты)
# ============================================
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


# ============================================
# СОСТОЯНИЕ: SALES (ввод продаж)
# ============================================
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


# ============================================
# СОСТОЯНИЕ: CHECKS (ввод чеков)
# ============================================
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


# ============================================
# СОСТОЯНИЕ: COST_PRICE (ввод себестоимости)
# ============================================
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


# ============================================
# СОСТОЯНИЕ: EXPENSES (ввод расходов)
# ============================================
async def process_expenses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка расходов"""
    try:
        expenses = float(update.message.text.replace(',', '.'))
        if expenses < 0:
            raise ValueError
        
        context.user_data['expenses'] = expenses
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


# ============================================
# СОСТОЯНИЕ: PHOTO (загрузка фото)
# ============================================
async def process_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Обработка фото"""
    await update.message.reply_text(PROCESSING_PHOTO)
    
    photo = update.message.photo[-1]
    file = await photo.get_file()
    photo_bytes = await file.download_as_bytearray()
    
    values = extract_from_photo(photo_bytes)
    
    if values and len(values) >= 4:
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
            logger.error(f"Ошибка распознавания фото: {e}")
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


# ============================================
# СОСТОЯНИЕ: CONFIRM (подтверждение)
# ============================================
async def confirm_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Подтверждение и сохранение"""
    query = update.callback_query
    await query.answer()
    
    data = context.user_data
    
    success = add_data_to_db(
        data['date'],
        data['day_of_week'],
        data['sales'],
        data['checks'],
        data['cost_price'],
        data['expenses']
    )
    
    if success:
        await safe_edit_message(
            query,
            f"{DATA_SAVED}\n\nЧто хотите сделать дальше?",
            reply_markup=main_menu_keyboard()
        )
    else:
        await safe_edit_message(
            query,
            "❌ Ошибка сохранения. Попробуйте ещё раз.",
            reply_markup=main_menu_keyboard()
        )
    
    context.user_data.clear()
    return MENU


async def edit_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Редактирование данных"""
    query = update.callback_query
    await query.answer()
    await safe_edit_message(
        query,
        EDIT_PROMPT,
        reply_markup=edit_keyboard()
    )
    return EDIT


async def cancel_from_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отмена из подтверждения"""
    query = update.callback_query
    await query.answer()
    context.user_data.clear()
    await safe_edit_message(
        query,
        CANCELLED,
        reply_markup=main_menu_keyboard()
    )
    return MENU


# ============================================
# СОСТОЯНИЕ: EDIT (редактирование)
# ============================================
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
    
    await safe_edit_message(
        query,
        messages[field],
        reply_markup=back_keyboard()
    )
    
    return states[field]


async def back_to_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Возврат к подтверждению"""
    query = update.callback_query
    await query.answer()
    
    text = f"""{CONFIRM_DATA}

📅 {DATE_LABEL}: {context.user_data.get('date', 'N/A')} ({context.user_data.get('day_of_week', '')})
💰 {SALES_LABEL}: {context.user_data.get('sales', 0):,.0f}
🧾 {CHECKS_LABEL}: {context.user_data.get('checks', 0)}
💸 {COST_PRICE_LABEL}: {context.user_data.get('cost_price', 0):,.0f}
📉 {EXPENSES_LABEL}: {context.user_data.get('expenses', 0):,.0f}
"""
    
    await safe_edit_message(
        query,
        text,
        reply_markup=confirm_keyboard()
    )
    return CONFIRM


# ============================================
# СОСТОЯНИЕ: TARGETS (установка целей)
# ============================================
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
        
        now = datetime.now()
        month = now.strftime('%B')
        year = now.year
        
        set_targets_db(sales_target, net_profit_target, expenses_target, month, year)
        
        await update.message.reply_text(
            f"{TARGETS_SET}\n\nЧто хотите сделать дальше?",
            reply_markup=main_menu_keyboard()
        )
        context.user_data.clear()
        return MENU
    except ValueError:
        await update.message.reply_text(
            f"{INVALID_FORMAT}\n\n{TARGETS_INSTRUCTION}",
            reply_markup=back_keyboard()
        )
        return TARGETS


# ============================================
# НАВИГАЦИЯ: Кнопка "Назад"
# ============================================
async def go_back(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Кнопка назад - возврат к выбору метода"""
    query = update.callback_query
    await query.answer()
    await safe_edit_message(
        query,
        SELECT_METHOD,
        reply_markup=method_keyboard()
    )
    return METHOD


async def back_to_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Возврат в главное меню"""
    query = update.callback_query
    await query.answer()
    context.user_data.clear()
    await safe_edit_message(
        query,
        MAIN_MENU_TEXT,
        reply_markup=main_menu_keyboard()
    )
    return MENU


# ============================================
# FALLBACK: /cancel и /start (сброс состояния)
# ============================================
async def cancel_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Отмена через /cancel"""
    context.user_data.clear()
    
    if update.callback_query:
        await update.callback_query.answer()
        await safe_edit_message(
            update.callback_query,
            CANCELLED,
            reply_markup=main_menu_keyboard()
        )
    else:
        await update.message.reply_text(
            CANCELLED,
            reply_markup=main_menu_keyboard()
        )
    return MENU


# ============================================
# НАСТРОЙКА ОБРАБОТЧИКОВ
# ============================================
def setup_handlers(application):
    """Регистрация всех обработчиков"""
    
    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler('start', start_entry)
        ],
        states={
            # ГЛАВНОЕ МЕНЮ
            MENU: [
                CallbackQueryHandler(menu_add_data, pattern='^add_data$'),
                CallbackQueryHandler(menu_view_report, pattern='^view_report$'),
                CallbackQueryHandler(menu_view_analytics, pattern='^analytics$'),
                CallbackQueryHandler(menu_set_targets, pattern='^set_targets$'),
            ],
            # ВЫБОР МЕТОДА
            METHOD: [
                CallbackQueryHandler(method_manual, pattern='^method_manual$'),
                CallbackQueryHandler(method_photo, pattern='^method_photo$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
            # ПОШАГОВЫЙ ВВОД
            DATE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_date),
                CallbackQueryHandler(go_back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
            SALES: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_sales),
                CallbackQueryHandler(go_back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
            CHECKS: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_checks),
                CallbackQueryHandler(go_back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
            COST_PRICE: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_cost_price),
                CallbackQueryHandler(go_back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
            EXPENSES: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_expenses),
                CallbackQueryHandler(go_back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
            # ФОТО
            PHOTO: [
                MessageHandler(filters.PHOTO, process_photo),
                CallbackQueryHandler(go_back, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
            # ПОДТВЕРЖДЕНИЕ
            CONFIRM: [
                CallbackQueryHandler(confirm_data, pattern='^confirm$'),
                CallbackQueryHandler(edit_data, pattern='^edit$'),
                CallbackQueryHandler(cancel_from_confirm, pattern='^cancel_add$'),
            ],
            # РЕДАКТИРОВАНИЕ
            EDIT: [
                CallbackQueryHandler(edit_field, pattern='^edit_'),
                CallbackQueryHandler(back_to_confirm, pattern='^back_to_confirm$'),
            ],
            # ЦЕЛИ
            TARGETS: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, process_targets),
                CallbackQueryHandler(back_to_main, pattern='^back$'),
                CallbackQueryHandler(back_to_main, pattern='^back_to_main$'),
            ],
        },
        fallbacks=[
            CommandHandler('cancel', cancel_command),
            CommandHandler('start', start_entry),  # <-- добавлен для сброса состояния
        ],
        per_message=False,
        name="pnl_bot_conversation",
        persistent=False,
    )
    
    application.add_handler(conv_handler)