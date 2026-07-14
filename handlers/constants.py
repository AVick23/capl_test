# Состояния для ConversationHandler
MENU, METHOD, DATE, SALES, CHECKS, COST_PRICE, EXPENSES, PHOTO, CONFIRM, EDIT, TARGETS = range(11)

# Текстовые сообщения
START_MESSAGE = (
    "👋 Добро пожаловать!\n\n"
    "Я помогу вам вести учёт ПНЛ:\n"
    "• Добавлять данные за день\n"
    "• Генерировать Excel-отчёты\n"
    "• Смотреть аналитику\n\n"
    "Выберите действие:"
)

MAIN_MENU_TEXT = "🏠 Главное меню"
ADD_DATA_TEXT = "➕ Добавить данные"
VIEW_REPORT_TEXT = "📊 Отчёт"
ANALYTICS_TEXT = "📈 Аналитика"
CANCEL_TEXT = "❌ Отмена"
BACK_TEXT = "🔙 Назад"
CONFIRM_TEXT = "✅ Подтвердить"
EDIT_TEXT = "✏️ Редактировать"
SAVE_TEXT = "💾 Сохранить"
TARGETS_TEXT = "🎯 Цели месяца"

# Сообщения о процессе
PROCESSING_PHOTO = "⏳ Обрабатываю фото..."
DATA_SAVED = "✅ Данные успешно сохранены!"
NO_DATA = "❌ Нет данных. Добавьте через меню."
INVALID_FORMAT = "❌ Неверный формат. Попробуйте ещё раз."
PHOTO_FAILED = "❌ Не удалось распознать. Введите данные вручную."
CANCELLED = "❌ Действие отменено"
WELCOME_BACK = "С возвращением! Выберите действие:"

# Сообщения для пошагового ввода
SELECT_METHOD = "📝 Как добавить данные?"
MANUAL_METHOD = "Ввести вручную"
PHOTO_METHOD = "Загрузить фото"

ENTER_DATE = "📅 Введите дату (ДД.ММ):\nПример: 15.05"
ENTER_SALES = "💰 Введите сумму продаж:"
ENTER_CHECKS = "🧾 Введите количество чеков:"
ENTER_COST_PRICE = "💸 Введите себестоимость:"
ENTER_EXPENSES = "📉 Введите расходы:"

CONFIRM_DATA = "🔍 Проверьте данные:"
EDIT_PROMPT = "✏️ Что изменить?"

DATE_LABEL = "Дата"
SALES_LABEL = "Продажи"
CHECKS_LABEL = "Чеки"
COST_PRICE_LABEL = "Себестоимость"
EXPENSES_LABEL = "Расходы"

REPORT_GENERATED = "📊 Отчёт готов!"
DASHBOARD_GENERATED = "📈 Аналитика готова!"

# Сообщения для целей
TARGETS_INSTRUCTION = (
    "🎯 Введите цели на месяц\n"
    "Формат: Продажи Прибыль Расходы\n\n"
    "Пример:\n"
    "500000 200000 300000"
)
TARGETS_SET = "✅ Цели установлены!"