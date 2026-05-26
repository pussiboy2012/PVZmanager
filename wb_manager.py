import pandas as pd
from models import db, IncomeTransaction
from datetime import datetime

def allowed_wb_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def import_wb_excel(filepath, sheet_name='Sheet1'):
    """
    Импорт отчёта Wildberries.
    Ожидается лист с колонками: 'ID ПВЗ', 'Дата', и множество числовых колонок с начислениями/удержаниями.
    Пропускаются строки, где 'ID ПВЗ' не является числом (например, 'Собственник').
    Для каждой даты создаются записи WbTransaction по каждому ненулевому значению в числовых колонках.
    """
    xl = pd.ExcelFile(filepath)
    if sheet_name not in xl.sheet_names:
        raise ValueError(f"Лист '{sheet_name}' не найден. Доступны: {xl.sheet_names}")

    df = xl.parse(sheet_name, dtype=str)  # читаем всё как строки для начала

    # Определим колонки, которые не нужно импортировать как услуги
    exclude_cols = {'ID ПВЗ', 'Регион', 'Адрес', 'Дата', 'Оборот', 'Оборот дорогостоя',
                    'Продажа по тарифу', 'Возврат по тарифу',  # эти могут быть нулевыми, но если нужны — можно оставить
                    'Продажа курьером', 'Возврат проданного курьером',
                    'Продажа курьерской доставкой', 'Возврат курьеской доставкой',
                    'Продажа дорогостоя товара', 'Возврат дорогостоя товара',
                    'Изменение баланса', 'Итог'}  # 'Итог' – суммарный итог, не нужно дублировать

    # Преобразуем даты
    df['datetime'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', errors='coerce')
    df = df.dropna(subset=['datetime'])

    # Оставляем только строки, где 'ID ПВЗ' можно преобразовать в число (исключаем 'Собственник' и пустые)
    df['pvz_id_num'] = pd.to_numeric(df['ID ПВЗ'], errors='coerce')
    df = df.dropna(subset=['pvz_id_num'])

    # Удаляем старые записи за те даты, которые есть в файле (чтобы не дублировать)
    dates_in_file = df['datetime'].dt.date.unique()
    for d in dates_in_file:
        start = datetime.combine(d, datetime.min.time())
        end = datetime.combine(d, datetime.max.time())
        IncomeTransaction.query.filter(IncomeTransaction.datetime.between(start, end)).delete()

    # Подготавливаем список для массовой вставки
    transactions = []

    # Перебираем строки
    for _, row in df.iterrows():
        pvz_id = int(row['pvz_id_num'])
        pvz_name = str(row.get('Название ПВЗ', '')) if pd.notna(row.get('Название ПВЗ')) else ''
        dt = row['datetime'].to_pydatetime()

        # Перебираем все колонки, кроме исключённых и уже обработанных
        for col in df.columns:
            if col in exclude_cols or col in ['datetime', 'pvz_id_num', 'Название ПВЗ', 'Дата', 'ID ПВЗ', 'Регион', 'Адрес']:
                continue
            # Пытаемся преобразовать значение в число
            try:
                val = float(row[col]) if pd.notna(row[col]) else 0.0
            except (ValueError, TypeError):
                continue
            if val == 0:
                continue
            # Создаём запись
            trans = IncomeTransaction(
                pvz_id=pvz_id,
                pvz_name=pvz_name,
                service=str(col),
                datetime=dt,
                service_cost=val
            )
            transactions.append(trans)

    # Массовое добавление
    db.session.add_all(transactions)
    db.session.commit()
    return len(transactions)