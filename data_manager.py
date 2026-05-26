import os
import pandas as pd
from models import db, Employee, Shift
from datetime import datetime, date
from config import Config

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS

def parse_time_interval(t_str):
    if not t_str or not isinstance(t_str, str):
        return None, None
    # Убираем пробелы
    t_str = t_str.strip()
    # Если нет дефиса, возможно, указано только одно время (тогда интервал не определён)
    if '-' not in t_str:
        # Пробуем распарсить как одиночное время (например, "18" или "18:00")
        try:
            # Если строка состоит только из цифр (часы без минут)
            if t_str.isdigit():
                start = f"{int(t_str):02d}:00"
                return start, None
            # Иначе пытаемся распарсить в формате HH:MM
            start = datetime.strptime(t_str, '%H:%M').time().strftime('%H:%M')
            return start, None
        except ValueError:
            return None, None
    # Если дефис есть, разбиваем
    start_str, end_str = t_str.split('-')
    start_str = start_str.strip()
    end_str = end_str.strip()
    # Пытаемся распарсить каждую часть
    try:
        # Если часть состоит только из цифр (часы), дополняем минутами
        if start_str.isdigit():
            start_str = f"{int(start_str):02d}:00"
        if end_str.isdigit():
            end_str = f"{int(end_str):02d}:00"
        start = datetime.strptime(start_str, '%H:%M').time().strftime('%H:%M')
        end = datetime.strptime(end_str, '%H:%M').time().strftime('%H:%M')
        return start, end
    except ValueError:
        return None, None

def import_excel(filepath):
    xl = pd.ExcelFile(filepath)
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        if 'Смена (дата)' not in df.columns or 'время работы' not in df.columns:
            continue
        emp = Employee.query.filter_by(name=sheet_name).first()
        if not emp:
            emp = Employee(name=sheet_name)
            db.session.add(emp)
            db.session.commit()

        for _, row in df.iterrows():
            date_val = row['Смена (дата)']
            if pd.isna(date_val):
                continue
            if isinstance(date_val, datetime):
                shift_date = date_val.date()
            else:
                try:
                    shift_date = datetime.strptime(str(date_val), '%Y-%m-%d').date()
                except:
                    continue

            start, end = parse_time_interval(str(row['время работы']))
            hours = float(row.get('Часов', 0) or 0)
            rate = float(row.get('Ставка', 0) or 0)
            bonus = float(row.get('Выплата заслуги', 0) or 0)
            deduction = float(row.get('Снятие вычет', 0) or 0)
            status = 'ДА' in str(row.get('Была смена ?', 'ДА'))

            shift = Shift(
                employee_id=emp.id,
                date=shift_date,
                start_time=start,
                end_time=end,
                hours=hours,
                rate=rate,
                bonus=bonus,
                deduction=deduction,
                status=status
            )
            db.session.add(shift)
    db.session.commit()