import pandas as pd
from models import db, IncomeTransaction
from datetime import datetime

def allowed_income_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

def import_income_excel(filepath):
    """
    Импортирует данные из листа 'Транзакции' Excel-файла.
    Очищает старые данные за тот же период (по датам из файла) и добавляет новые.
    """
    xl = pd.ExcelFile(filepath)
    if 'Транзакции' not in xl.sheet_names:
        raise ValueError("Лист 'Транзакции' не найден в файле")

    df = xl.parse('Транзакции')
    required_cols = ['Время (мск)', 'Услуга', 'Стоимость услуги, руб']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"В листе 'Транзакции' отсутствует колонка: {col}")

    # Преобразуем даты
    df['datetime'] = pd.to_datetime(df['Время (мск)'], errors='coerce')
    df = df.dropna(subset=['datetime'])
    dates_in_file = df['datetime'].dt.date.unique()

    # Удаляем старые записи за эти даты (чтобы не дублировать)
    for d in dates_in_file:
        start = datetime.combine(d, datetime.min.time())
        end = datetime.combine(d, datetime.max.time())
        IncomeTransaction.query.filter(IncomeTransaction.datetime.between(start, end)).delete()

    # Заполняем недостающие поля
    df['pvz_id'] = df.get('ID ПВЗ', 0).fillna(0).astype(int)
    df['pvz_name'] = df.get('Название ПВЗ', '').fillna('')
    df['service'] = df['Услуга'].fillna('')
    df['order_ref'] = df.get('Заказ/Возврат/Дата', '').fillna('')
    df['product'] = df.get('Товар', '').fillna('')
    df['cost_per_item'] = pd.to_numeric(df.get('Стоимость заказа/товара за 1шт, руб', 0), errors='coerce').fillna(0)
    df['cost_without_vat'] = pd.to_numeric(df.get('Стоимость заказа/товара за 1шт без НДС, руб', 0), errors='coerce').fillna(0)
    df['quantity'] = pd.to_numeric(df.get('Количество', 1), errors='coerce').fillna(1)
    df['payment_type'] = df.get('Тип оплаты', '').fillna('')
    df['service_cost'] = pd.to_numeric(df['Стоимость услуги, руб'], errors='coerce').fillna(0)
    df['tariff'] = pd.to_numeric(df.get('Тариф', 0), errors='coerce').fillna(0)
    df['unit'] = df.get('Единица измерения', '').fillna('')
    df['branded'] = df.get('Брендированный', '').astype(str).str.lower().isin(['да', 'true', '1', 'yes'])
    df['tariff_applies'] = df.get('Применение тарифа', '').fillna('')
    df['region'] = df.get('Регион (бренд)', '').fillna('')
    df['tariff_zone'] = df.get('Тарифная зона (бренд)', '').fillna('')

    # Массовая вставка
    transactions = []
    for _, row in df.iterrows():
        trans = IncomeTransaction(
            pvz_id=int(row['pvz_id']),
            pvz_name=str(row['pvz_name']),
            service=str(row['service']),
            datetime=row['datetime'].to_pydatetime(),
            order_ref=str(row['order_ref']),
            product=str(row['product']),
            cost_per_item=float(row['cost_per_item']),
            cost_without_vat=float(row['cost_without_vat']),
            quantity=float(row['quantity']),
            payment_type=str(row['payment_type']),
            service_cost=float(row['service_cost']),
            tariff=float(row['tariff']),
            unit=str(row['unit']),
            branded=bool(row['branded']),
            tariff_applies=str(row['tariff_applies']),
            region=str(row['region']),
            tariff_zone=str(row['tariff_zone'])
        )
        transactions.append(trans)

    db.session.add_all(transactions)
    db.session.commit()
    return len(transactions)


def import_wildberries_excel(filepath):
    """
    Импортирует данные из листа Sheet1 Excel-файла Wildberries.
    Фильтрует строки с "Собственник" и пустые даты.
    Каждая строка = одна транзакция с суммой "Итог".
    """
    xl = pd.ExcelFile(filepath)
    if 'Sheet1' not in xl.sheet_names:
        raise ValueError("Лист 'Sheet1' не найден в файле")
    df = xl.parse('Sheet1')

    # Определяем колонку "Итог" (последняя)
    total_col = df.columns[-1]  # предположительно "Итог"

    # Преобразуем даты
    df['date'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', errors='coerce')
    df = df.dropna(subset=['date'])

    # Фильтруем: исключаем строки, где ID ПВЗ не число или равно "Собственник"
    df = df[df['ID ПВЗ'].apply(lambda x: str(x).replace('.0', '').isdigit())]
    df = df[~df['ID ПВЗ'].astype(str).str.contains('Собственник', na=False)]

    # Оставляем только нужные колонки
    df = df[['ID ПВЗ', 'Регион', 'Адрес', 'date', total_col]]
    df = df.rename(columns={'ID ПВЗ': 'pvz_id', 'Регион': 'region', 'Адрес': 'address', total_col: 'total'})

    # Удаляем дубликаты по дате (если несколько строк за день – суммируем? Но в файле по одной строке на день)
    # Группируем на случай дублей
    df = df.groupby('date').agg(
        {'total': 'sum', 'pvz_id': 'first', 'region': 'first', 'address': 'first'}).reset_index()

    # Удаляем старые записи Wildberries за эти даты
    dates_in_file = df['date'].dt.date.unique()
    for d in dates_in_file:
        start = datetime.combine(d, datetime.min.time())
        end = datetime.combine(d, datetime.max.time())
        IncomeTransaction.query.filter(
            IncomeTransaction.datetime.between(start, end),
            IncomeTransaction.source == 'wildberries'
        ).delete()

    # Добавляем новые транзакции
    transactions = []
    for _, row in df.iterrows():
        trans = IncomeTransaction(
            pvz_id=int(row['pvz_id']),
            pvz_name=str(row['address']),
            service='Итог дня (WB)',
            datetime=row['date'].to_pydatetime(),
            order_ref='',
            product='',
            cost_per_item=0,
            cost_without_vat=0,
            quantity=1,
            payment_type='',
            service_cost=float(row['total']),
            tariff=0,
            unit='руб',
            branded=False,
            tariff_applies='',
            region=str(row['region']),
            tariff_zone='',
            source='wildberries'
        )
        transactions.append(trans)

    db.session.add_all(transactions)
    db.session.commit()
    return len(transactions)