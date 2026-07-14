import sqlite3
import pytesseract
from PIL import Image
import io
import re
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Для серверного режима без GUI
from config import DB_PATH
from .constants import *


def init_db():
    """Инициализация базы данных"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS daily_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL UNIQUE,
            day_of_week TEXT,
            sales REAL NOT NULL,
            checks INTEGER NOT NULL,
            cost_price REAL NOT NULL,
            expenses REAL NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS targets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_target REAL,
            net_profit_target REAL,
            expenses_target REAL,
            month TEXT,
            year INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    conn.commit()
    conn.close()


def add_data_to_db(date, day_of_week, sales, checks, cost_price, expenses):
    """Добавление данных в БД"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
            INSERT OR REPLACE INTO daily_data 
            (date, day_of_week, sales, checks, cost_price, expenses)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (date, day_of_week, sales, checks, cost_price, expenses))
        conn.commit()
        success = True
    except Exception as e:
        print(f"Error adding data: {e}")
        success = False
    finally:
        conn.close()
    
    return success


def get_all_data():
    """Получение всех данных"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT date, day_of_week, sales, checks, cost_price, expenses 
        FROM daily_data 
        ORDER BY date
    """)
    rows = cursor.fetchall()
    conn.close()
    return rows


def delete_data(date):
    """Удаление данных за конкретную дату"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM daily_data WHERE date = ?", (date,))
    conn.commit()
    conn.close()


def set_targets_db(sales_target, net_profit_target, expenses_target, month, year):
    """Установка целей"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("""
        DELETE FROM targets 
        WHERE month = ? AND year = ?
    """, (month, year))
    
    cursor.execute("""
        INSERT INTO targets (sales_target, net_profit_target, expenses_target, month, year)
        VALUES (?, ?, ?, ?, ?)
    """, (sales_target, net_profit_target, expenses_target, month, year))
    
    conn.commit()
    conn.close()


def get_targets_db(month, year):
    """Получение целей"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT sales_target, net_profit_target, expenses_target
        FROM targets
        WHERE month = ? AND year = ?
        ORDER BY created_at DESC
        LIMIT 1
    """, (month, year))
    result = cursor.fetchone()
    conn.close()
    return result


def parse_date(date_str):
    """Парсинг даты и определение дня недели"""
    days_ru = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']
    
    try:
        # Пробуем разные форматы
        for fmt in ['%d.%m', '%d.%m.%Y', '%Y-%m-%d']:
            try:
                dt = datetime.strptime(date_str, fmt)
                if fmt == '%d.%m':
                    dt = dt.replace(year=datetime.now().year)
                day_of_week = days_ru[dt.weekday()]
                return dt.strftime('%d.%m'), day_of_week
            except ValueError:
                continue
        return None, None
    except:
        return None, None


def extract_from_photo(photo_bytes):
    """Извлечение данных из фото через OCR"""
    try:
        # Открываем изображение
        image = Image.open(io.BytesIO(photo_bytes))
        
        # OCR с русским языком
        text = pytesseract.image_to_string(image, lang='rus')
        
        # Пытаемся извлечь числа из текста
        numbers = re.findall(r'\d+(?:[.,]\d+)?', text)
        
        # Преобразуем в float
        values = []
        for num in numbers:
            try:
                values.append(float(num.replace(',', '.')))
            except:
                pass
        
        # Возвращаем первые 4-5 найденных значений
        return values[:5] if len(values) >= 4 else None
        
    except Exception as e:
        print(f"OCR Error: {e}")
        return None


def generate_excel_report():
    """Генерация Excel-отчёта"""
    data = get_all_data()
    
    if not data:
        return None
    
    # Создаем DataFrame
    df = pd.DataFrame(data, columns=[
        'Дата', 'День недели', 'Продажи', 'Чеки', 'Себестоимость', 'Расходы'
    ])
    
    # Добавляем расчетные поля
    df['Ср. чек'] = df['Продажи'] / df['Чеки']
    df['Валовая прибыль'] = df['Продажи'] - df['Себестоимость']
    df['Чистая прибыль'] = df['Валовая прибыль'] - df['Расходы']
    df['Себестоимость %'] = (df['Себестоимость'] / df['Продажи'] * 100).round(1)
    df['Валовая %'] = (df['Валовая прибыль'] / df['Продажи'] * 100).round(1)
    
    # Сохраняем в Excel
    filename = 'pnl_report.xlsx'
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ПНЛ', index=False)
        
        # Форматирование
        workbook = writer.book
        worksheet = writer.sheets['ПНЛ']
        
        # Настройка колонок
        worksheet.column_dimensions['A'].width = 12  # Дата
        worksheet.column_dimensions['B'].width = 12  # День недели
        worksheet.column_dimensions['C'].width = 15  # Продажи
        worksheet.column_dimensions['D'].width = 10  # Чеки
        worksheet.column_dimensions['E'].width = 18  # Себестоимость
        worksheet.column_dimensions['F'].width = 15  # Расходы
        worksheet.column_dimensions['G'].width = 12  # Ср. чек
        worksheet.column_dimensions['H'].width = 18  # Валовая прибыль
        worksheet.column_dimensions['I'].width = 18  # Чистая прибыль
        worksheet.column_dimensions['J'].width = 15  # Себестоимость %
        worksheet.column_dimensions['K'].width = 12  # Валовая %
        
        # Заголовки жирным
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            cell.font = Font(bold=True, color='FFFFFF')
        
        # Добавляем итоги
        last_row = len(df) + 1
        worksheet.cell(row=last_row + 1, column=1, value='ИТОГО:')
        worksheet.cell(row=last_row + 1, column=1).font = Font(bold=True)
        
        worksheet.cell(row=last_row + 1, column=3, value=f'=SUM(C2:C{last_row})')
        worksheet.cell(row=last_row + 1, column=4, value=f'=SUM(D2:D{last_row})')
        worksheet.cell(row=last_row + 1, column=5, value=f'=SUM(E2:E{last_row})')
        worksheet.cell(row=last_row + 1, column=6, value=f'=SUM(F2:F{last_row})')
        worksheet.cell(row=last_row + 1, column=8, value=f'=SUM(H2:H{last_row})')
        worksheet.cell(row=last_row + 1, column=9, value=f'=SUM(I2:I{last_row})')
    
    return filename


def generate_dashboard():
    """Генерация дашборда"""
    data = get_all_data()
    
    if not data:
        return None, None
    
    df = pd.DataFrame(data, columns=[
        'date', 'day', 'sales', 'checks', 'cost', 'expenses'
    ])
    
    # Расчёты
    df['gross_profit'] = df['sales'] - df['cost']
    df['net_profit'] = df['gross_profit'] - df['expenses']
    df['avg_check'] = df['sales'] / df['checks']
    
    # Создаём графики
    fig, axes = plt.subplots(2, 2, figsize=(16, 10))
    fig.suptitle('Аналитика ПНЛ', fontsize=20, fontweight='bold')
    
    # 1. Продажи по дням
    ax1 = axes[0, 0]
    ax1.bar(df['date'], df['sales'], color='#4472C4', alpha=0.8)
    ax1.set_title('Продажи по дням', fontsize=14, fontweight='bold')
    ax1.set_xlabel('Дата')
    ax1.set_ylabel('Сумма')
    ax1.tick_params(axis='x', rotation=45)
    ax1.grid(axis='y', alpha=0.3)
    
    # 2. Чистая прибыль
    ax2 = axes[0, 1]
    colors = ['#2ecc71' if p >= 0 else '#e74c3c' for p in df['net_profit']]
    ax2.bar(df['date'], df['net_profit'], color=colors, alpha=0.8)
    ax2.axhline(y=0, color='black', linewidth=1, linestyle='--')
    ax2.set_title('Чистая прибыль', fontsize=14, fontweight='bold')
    ax2.set_xlabel('Дата')
    ax2.set_ylabel('Сумма')
    ax2.tick_params(axis='x', rotation=45)
    ax2.grid(axis='y', alpha=0.3)
    
    # 3. Средний чек
    ax3 = axes[1, 0]
    ax3.plot(df['date'], df['avg_check'], marker='o', linewidth=2, color='#f39c12')
    ax3.fill_between(df['date'], df['avg_check'], alpha=0.3, color='#f39c12')
    ax3.set_title('Средний чек', fontsize=14, fontweight='bold')
    ax3.set_xlabel('Дата')
    ax3.set_ylabel('Сумма')
    ax3.tick_params(axis='x', rotation=45)
    ax3.grid(alpha=0.3)
    
    # 4. Структура расходов
    ax4 = axes[1, 1]
    total_cost = df['cost'].sum()
    total_expenses = df['expenses'].sum()
    total_profit = df['net_profit'].sum()
    
    if total_profit > 0:
        sizes = [total_cost, total_expenses, total_profit]
        labels = ['Себестоимость', 'Расходы', 'Прибыль']
        colors_pie = ['#3498db', '#e74c3c', '#2ecc71']
    else:
        sizes = [total_cost, total_expenses]
        labels = ['Себестоимость', 'Расходы']
        colors_pie = ['#3498db', '#e74c3c']
    
    ax4.pie(sizes, labels=labels, colors=colors_pie, autopct='%1.1f%%', startangle=90)
    ax4.set_title('Структура', fontsize=14, fontweight='bold')
    
    plt.tight_layout()
    
    # Сохраняем
    filename = 'dashboard.png'
    plt.savefig(filename, dpi=150, bbox_inches='tight')
    plt.close()
    
    # Генерируем текстовую аналитику
    analytics = generate_analytics_text(df)
    
    return filename, analytics


def generate_analytics_text(df):
    """Генерация текстовой аналитики"""
    total_sales = df['sales'].sum()
    total_profit = df['net_profit'].sum()
    total_expenses = df['expenses'].sum()
    total_cost = df['cost'].sum()
    days = len(df)
    
    avg_daily_sales = total_sales / days
    avg_daily_profit = total_profit / days
    
    text = "📊 АНАЛИТИКА\n\n"
    
    # Основные показатели
    text += f"💰 Продажи: {total_sales:,.0f}\n"
    text += f"💸 Себестоимость: {total_cost:,.0f} ({total_cost/total_sales*100:.1f}%)\n"
    text += f"📉 Расходы: {total_expenses:,.0f} ({total_expenses/total_sales*100:.1f}%)\n"
    text += f"📈 Чистая прибыль: {total_profit:,.0f}\n\n"
    
    # Что хорошо
    good = []
    if total_profit > 0:
        good.append(f"✅ Прибыль положительная: {total_profit:,.0f}")
    if avg_daily_profit > 10000:
        good.append(f"✅ Средний день: {avg_daily_profit:,.0f}")
    
    text += "🟢 ХОРОШО:\n"
    text += "\n".join(good) if good else "—暂无\n"
    
    # Что плохо
    bad = []
    if total_profit < 0:
        bad.append(f"❌ Убыток: {total_profit:,.0f}")
    if total_expenses / total_sales > 0.5:
        bad.append(f"❌ Высокие расходы: {total_expenses/total_sales*100:.1f}%")
    if total_cost / total_sales > 0.4:
        bad.append(f"❌ Высокая себестоимость: {total_cost/total_sales*100:.1f}%")
    
    text += "\n🔴 ПЛОХО:\n"
    text += "\n".join(bad) if bad else "— нет замечаний\n"
    
    # Средние показатели
    text += f"\n📊 СРЕДНИЕ:\n"
    text += f"• Продажи/день: {avg_daily_sales:,.0f}\n"
    text += f"• Прибыль/день: {avg_daily_profit:,.0f}\n"
    text += f"• Средний чек: {df['avg_check'].mean():,.0f}\n"
    
    return text