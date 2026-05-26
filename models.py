from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime, date

db = SQLAlchemy()

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(128), nullable=False)

class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)

class Shift(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'))
    date = db.Column(db.Date, nullable=False)
    start_time = db.Column(db.String(10))
    end_time = db.Column(db.String(10))
    hours = db.Column(db.Float, default=0)          # <-- добавлено default
    rate = db.Column(db.Float, default=0)           # <--
    bonus = db.Column(db.Float, default=0)          # <--
    deduction = db.Column(db.Float, default=0)      # <--
    status = db.Column(db.Boolean, default=True)

    employee = db.relationship('Employee', backref='shifts')

# Добавить в models.py после класса Shift

class IncomeTransaction(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    pvz_id = db.Column(db.Integer)
    pvz_name = db.Column(db.String(200))
    service = db.Column(db.String(200))
    datetime = db.Column(db.DateTime, nullable=False)
    order_ref = db.Column(db.String(100))
    product = db.Column(db.String(100))
    cost_per_item = db.Column(db.Float)
    cost_without_vat = db.Column(db.Float)
    quantity = db.Column(db.Float, default=1)
    payment_type = db.Column(db.String(50))
    service_cost = db.Column(db.Float, default=0)
    tariff = db.Column(db.Float)
    unit = db.Column(db.String(20))
    branded = db.Column(db.Boolean, default=True)
    tariff_applies = db.Column(db.String(50))
    region = db.Column(db.String(100))
    tariff_zone = db.Column(db.String(100))
    source = db.Column(db.String(20), default='yandex')  # 'yandex' или 'wildberries'

class Expense(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(50), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    date = db.Column(db.Date, nullable=False, default=date.today)
    description = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Категории расходов (для удобства)
    @staticmethod
    def get_categories():
        return [
            'Интернет',
            'Электричество',
            'Вода',
            'Отопление',
            'Экопром (мусор)',
            'Содержание и ремонт',
            'Капремонт',
            'Охрана',
            'Камеры видеонаблюдения',
            'Вода питьевая',
            'Канцелярия',
            'Хозтовары',
            'Прочее'
        ]