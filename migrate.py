#!/usr/bin/env python3
"""
Скрипт для заполнения базы данных тестовыми данными.

Использование:
    python migrate.py                    # Добавить тестовые данные
    python migrate.py --clean            # Очистить БД и добавить данные
    python migrate.py --days 15          # Указать количество дней
    python migrate.py --help             # Справка
"""

import sqlite3
import random
from datetime import datetime, timedelta
import sys
import argparse
from config import DB_PATH
from handlers.utils import init_db


# ============================================
# НАСТРОЙКИ (можно менять под себя)
# ============================================
CONFIG = {
    # Период данных
    'start_date': datetime(2026, 7, 1),      # С какой даты начинаем
    'default_days': 14,                       # Количество дней по умолчанию
    
    # Диапазоны значений (в рублях)
    'sales_range': (8000, 25000),             # Продажи
    'checks_range': (15, 45),                 # Количество чеков
    'cost_price_percent': (0.30, 0.45),       # Себестоимость в % от продаж
    'expenses_range': (12000, 20000),         # Расходы (фиксированные)
    
    # Множители для выходных (сб, вс - обычно выше продажи)
    'weekend_multiplier': 1.4,
    
    # Цели на месяц
    'targets': {
        'sales_target': 500000,
        'net_profit_target': 150000,
        'expenses_target': 450000
    }
}

# Дни недели на русском
DAYS_RU = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']


def clear_database():
    """Очистка всех таблиц в базе данных"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        cursor.execute("DELETE FROM daily_data")
        cursor.execute("DELETE FROM targets")
        conn.commit()
        print("✅ База данных очищена")
        return True
    except Exception as e:
        print(f"❌ Ошибка очистки БД: {e}")
        return False
    finally:
        conn.close()


def check_existing_data():
    """Проверка наличия данных в БД"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        cursor.execute("SELECT COUNT(*) FROM daily_data")
        count = cursor.fetchone()[0]
        return count
    finally:
        conn.close()


def generate_daily_data(date, day_of_week):
    """Генерация данных для одного дня"""
    is_weekend = day_of_week in ['сб', 'вс']
    
    # Продажи
    sales_min, sales_max = CONFIG['sales_range']
    sales = random.randint(sales_min, sales_max)
    if is_weekend:
        sales = int(sales * CONFIG['weekend_multiplier'])
    
    # Чеки (пропорционально продажам с небольшим разбросом)
    checks_min, checks_max = CONFIG['checks_range']
    checks = random.randint(checks_min, checks_max)
    if is_weekend:
        checks = int(checks * CONFIG['weekend_multiplier'])
    
    # Себестоимость (процент от продаж)
    cost_min, cost_max = CONFIG['cost_price_percent']
    cost_percent = random.uniform(cost_min, cost_max)
    cost_price = sales * cost_percent
    
    # Расходы
    exp_min, exp_max = CONFIG['expenses_range']
    expenses = random.randint(exp_min, exp_max)
    
    return {
        'date': date.strftime('%d.%m'),
        'day_of_week': day_of_week,
        'sales': round(sales, 2),
        'checks': checks,
        'cost_price': round(cost_price, 2),
        'expenses': round(expenses, 2)
    }


def add_data_to_db(data_list):
    """Добавление списка данных в БД"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    success_count = 0
    error_count = 0
    
    for data in data_list:
        try:
            cursor.execute("""
                INSERT OR REPLACE INTO daily_data 
                (date, day_of_week, sales, checks, cost_price, expenses)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                data['date'],
                data['day_of_week'],
                data['sales'],
                data['checks'],
                data['cost_price'],
                data['expenses']
            ))
            success_count += 1
        except Exception as e:
            print(f"⚠️  Ошибка добавления {data['date']}: {e}")
            error_count += 1
    
    conn.commit()
    conn.close()
    
    return success_count, error_count


def set_monthly_targets():
    """Установка целей на текущий месяц"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    now = datetime.now()
    month = now.strftime('%B')
    year = now.year
    
    try:
        # Удаляем старые цели для этого месяца
        cursor.execute("""
            DELETE FROM targets 
            WHERE month = ? AND year = ?
        """, (month, year))
        
        # Добавляем новые цели
        cursor.execute("""
            INSERT INTO targets (sales_target, net_profit_target, expenses_target, month, year)
            VALUES (?, ?, ?, ?, ?)
        """, (
            CONFIG['targets']['sales_target'],
            CONFIG['targets']['net_profit_target'],
            CONFIG['targets']['expenses_target'],
            month,
            year
        ))
        
        conn.commit()
        print(f"✅ Цели на {month} {year} установлены")
        return True
    except Exception as e:
        print(f"❌ Ошибка установки целей: {e}")
        return False
    finally:
        conn.close()


def print_statistics(data_list):
    """Вывод статистики по сгенерированным данным"""
    if not data_list:
        return
    
    total_sales = sum(d['sales'] for d in data_list)
    total_checks = sum(d['checks'] for d in data_list)
    total_cost = sum(d['cost_price'] for d in data_list)
    total_expenses = sum(d['expenses'] for d in data_list)
    total_profit = total_sales - total_cost - total_expenses
    
    avg_daily_sales = total_sales / len(data_list)
    avg_check = total_sales / total_checks if total_checks > 0 else 0
    
    print("\n" + "=" * 60)
    print("📊 СТАТИСТИКА ТЕСТОВЫХ ДАННЫХ")
    print("=" * 60)
    print(f"📅 Дней: {len(data_list)}")
    print(f"💰 Продажи: {total_sales:,.0f} ₽")
    print(f"🧾 Чеков: {total_checks}")
    print(f"💸 Себестоимость: {total_cost:,.0f} ₽")
    print(f"📉 Расходы: {total_expenses:,.0f} ₽")
    print(f"📈 Чистая прибыль: {total_profit:,.0f} ₽")
    print("-" * 60)
    print(f"📊 Ср. продажи/день: {avg_daily_sales:,.0f} ₽")
    print(f"💵 Ср. чек: {avg_check:,.0f} ₽")
    print(f"📊 Себестоимость: {total_cost/total_sales*100:.1f}% от продаж")
    print("=" * 60)


def main():
    """Основная функция"""
    # Парсинг аргументов командной строки
    parser = argparse.ArgumentParser(
        description='Заполнение базы данных тестовыми данными для ПНЛ бота'
    )
    parser.add_argument(
        '--clean',
        action='store_true',
        help='Очистить базу данных перед заполнением'
    )
    parser.add_argument(
        '--days',
        type=int,
        default=CONFIG['default_days'],
        help=f'Количество дней для генерации (по умолчанию: {CONFIG["default_days"]})'
    )
    parser.add_argument(
        '--force',
        action='store_true',
        help='Заполнить БД даже если в ней уже есть данные'
    )
    
    args = parser.parse_args()
    
    print("\n" + "=" * 60)
    print("🚀 ГЕНЕРАЦИЯ ТЕСТОВЫХ ДАННЫХ")
    print("=" * 60)
    
    # Инициализация БД
    init_db()
    
    # Проверка существующих данных
    existing_count = check_existing_data()
    
    if existing_count > 0 and not args.clean and not args.force:
        print(f"\n⚠️  В базе уже есть {existing_count} записей!")
        print("Используйте один из вариантов:")
        print("  • python migrate.py --clean    (очистить и заполнить заново)")
        print("  • python migrate.py --force    (добавить к существующим)")
        print("  • python migrate.py --help     (показать справку)")
        sys.exit(1)
    
    # Очистка если нужно
    if args.clean:
        if not clear_database():
            sys.exit(1)
    
    # Генерация данных
    print(f"\n📅 Генерируем данные за {args.days} дней...")
    print(f"   Начало: {CONFIG['start_date'].strftime('%d.%m.%Y')}")
    
    data_list = []
    current_date = CONFIG['start_date']
    
    for i in range(args.days):
        day_of_week = DAYS_RU[current_date.weekday()]
        data = generate_daily_data(current_date, day_of_week)
        data_list.append(data)
        current_date += timedelta(days=1)
    
    # Добавление в БД
    print(f"\n💾 Сохраняем в базу данных...")
    success, errors = add_data_to_db(data_list)
    
    print(f"✅ Успешно добавлено: {success} записей")
    if errors > 0:
        print(f"⚠️  Ошибок: {errors}")
    
    # Установка целей
    set_monthly_targets()
    
    # Статистика
    print_statistics(data_list)
    
    print("\n✨ Готово! Теперь можете запустить бота и проверить данные.\n")


if __name__ == "__main__":
    main()