from telegram import InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from .constants import *


def main_menu_keyboard():
    """Главное меню"""
    keyboard = [
        [InlineKeyboardButton(ADD_DATA_TEXT, callback_data="add_data")],
        [InlineKeyboardButton(VIEW_REPORT_TEXT, callback_data="view_report")],
        [InlineKeyboardButton(ANALYTICS_TEXT, callback_data="analytics")],
        [InlineKeyboardButton(TARGETS_TEXT, callback_data="set_targets")]
    ]
    return InlineKeyboardMarkup(keyboard)


def method_keyboard():
    """Выбор метода добавления"""
    keyboard = [
        [InlineKeyboardButton(MANUAL_METHOD, callback_data="method_manual")],
        [InlineKeyboardButton(PHOTO_METHOD, callback_data="method_photo")],
        [InlineKeyboardButton(BACK_TEXT, callback_data="back_to_main")]
    ]
    return InlineKeyboardMarkup(keyboard)


def back_keyboard():
    """Кнопка назад"""
    keyboard = [[InlineKeyboardButton(BACK_TEXT, callback_data="back")]]
    return InlineKeyboardMarkup(keyboard)


def back_to_main_keyboard():
    """Вернуться в главное меню"""
    keyboard = [[InlineKeyboardButton(MAIN_MENU_TEXT, callback_data="main_menu")]]
    return InlineKeyboardMarkup(keyboard)


def confirm_keyboard():
    """Подтверждение данных"""
    keyboard = [
        [InlineKeyboardButton(CONFIRM_TEXT, callback_data="confirm")],
        [InlineKeyboardButton(EDIT_TEXT, callback_data="edit")],
        [InlineKeyboardButton(CANCEL_TEXT, callback_data="cancel_add")]
    ]
    return InlineKeyboardMarkup(keyboard)


def edit_keyboard():
    """Редактирование полей"""
    keyboard = [
        [InlineKeyboardButton(DATE_LABEL, callback_data="edit_date")],
        [InlineKeyboardButton(SALES_LABEL, callback_data="edit_sales")],
        [InlineKeyboardButton(CHECKS_LABEL, callback_data="edit_checks")],
        [InlineKeyboardButton(COST_PRICE_LABEL, callback_data="edit_cost_price")],
        [InlineKeyboardButton(EXPENSES_LABEL, callback_data="edit_expenses")],
        [InlineKeyboardButton(BACK_TEXT, callback_data="back_to_confirm")]
    ]
    return InlineKeyboardMarkup(keyboard)


def reply_keyboard():
    """Постоянная клавиатура (опционально)"""
    keyboard = [
        [KeyboardButton(MAIN_MENU_TEXT)]
    ]
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)