import os
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from sqlalchemy import func
from werkzeug.security import generate_password_hash, check_password_hash
from config import Config
from models import db, User, Employee, Shift, IncomeTransaction, Expense
from data_manager import allowed_file, import_excel
from flask import Flask, render_template, redirect, url_for, request, flash, jsonify
import calendar as cal
from datetime import datetime, date, timedelta
from income_manager import import_income_excel, allowed_income_file, import_wildberries_excel
from io import BytesIO
from flask import send_file
import pandas as pd

app = Flask(__name__)
app.config.from_object(Config)

db.init_app(app)

login_manager = LoginManager()
login_manager.login_view = 'login'
login_manager.init_app(app)


from flask_cors import CORS

# после создания app
CORS(app, resources={r"/api/*": {"origins": "*", "supports_credentials": True}})

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))

@app.context_processor
def utility_processor():
    return {'now': datetime.now}


from num2words import num2words

@app.template_filter('word_num')
def word_num_filter(number):
    if number is None:
        return "ноль рублей 00 копеек"
    rub = int(number)
    kop = int(round((number - rub) * 100))
    rub_str = num2words(rub, lang='ru')
    # Склоняем "рубль"
    if rub % 10 == 1 and rub % 100 != 11:
        rub_word = "рубль"
    elif 2 <= rub % 10 <= 4 and not (12 <= rub % 100 <= 14):
        rub_word = "рубля"
    else:
        rub_word = "рублей"
    return f"{rub_str} {rub_word} {kop:02d} копеек"


# Создание таблиц и администратора при запуске (замена before_first_request)
with app.app_context():
    db.create_all()
    if not User.query.filter_by(username="admin").first():
        user = User(username="admin", password=generate_password_hash("12345"))
        db.session.add(user)
        db.session.commit()

# ------------------- Основные маршруты -------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        user = User.query.filter_by(username=username).first()
        if user and check_password_hash(user.password, password):
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Неверные данные")
    return render_template("login.html")

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))

@app.route("/", methods=["GET", "POST"])
@login_required
def dashboard():
    # Получаем выбранный месяц/год для KPI и топов (по умолчанию текущий)
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    today = datetime.today()
    if not year:
        year = today.year
    if not month:
        month = today.month

    # Навигация по месяцам для KPI
    if month == 1:
        prev_year, prev_month = year - 1, 12
    else:
        prev_year, prev_month = year, month - 1
    if month == 12:
        next_year, next_month = year + 1, 1
    else:
        next_year, next_month = year, month + 1

    # --- KPI за выбранный месяц ---
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year + 1, 1, 1)
    else:
        end_date = date(year, month + 1, 1)

    total_employees = Employee.query.count()
    shifts_month = Shift.query.filter(Shift.date >= start_date, Shift.date < end_date).all()
    total_shifts = len(shifts_month)
    total_hours = sum((s.hours or 0) for s in shifts_month)

    start_datetime = datetime(year, month, 1)
    end_datetime = datetime(year + 1, 1, 1) if month == 12 else datetime(year, month + 1, 1)
    income_total = db.session.query(func.sum(IncomeTransaction.service_cost)).filter(
        IncomeTransaction.datetime >= start_datetime,
        IncomeTransaction.datetime < end_datetime
    ).scalar() or 0



    # --- ЕЖЕМЕСЯЧНЫЕ ДАННЫЕ ЗА ТЕКУЩИЙ ГОД (для графиков) ---
    months = []
    monthly_hours = []
    monthly_income = []
    for m in range(1, 13):
        start_of_month = date(year, m, 1)
        if m == 12:
            end_of_month = date(year + 1, 1, 1)
        else:
            end_of_month = date(year, m + 1, 1)
        # Часы
        shifts_m = Shift.query.filter(Shift.date >= start_of_month, Shift.date < end_of_month).all()
        hours_m = sum((s.hours or 0) for s in shifts_m)
        monthly_hours.append(round(hours_m, 1))
        # Доход
        start_dt = datetime(year, m, 1)
        end_dt = datetime(year + 1, 1, 1) if m == 12 else datetime(year, m + 1, 1)
        income_m = db.session.query(func.sum(IncomeTransaction.service_cost)).filter(
            IncomeTransaction.datetime >= start_dt,
            IncomeTransaction.datetime < end_dt
        ).scalar() or 0
        monthly_income.append(round(income_m, 2))
        months.append(cal.month_abbr[m])   # янв, фев, ...

    # --- Топ сотрудников по часам за выбранный месяц ---
    from collections import defaultdict
    employee_hours = defaultdict(float)
    for s in shifts_month:
        employee_hours[s.employee.name] += (s.hours or 0)
    top_employees = sorted(employee_hours.items(), key=lambda x: x[1], reverse=True)[:5]

    # --- Топ услуг по доходу за выбранный месяц ---
    service_income = db.session.query(
        IncomeTransaction.service,
        func.sum(IncomeTransaction.service_cost)
    ).filter(
        IncomeTransaction.datetime >= start_datetime,
        IncomeTransaction.datetime < end_datetime
    ).group_by(IncomeTransaction.service).all()
    top_services = sorted(service_income, key=lambda x: x[1], reverse=True)[:5]

    # --- Последние 5 смен ---
    recent_shifts = Shift.query.order_by(Shift.date.desc()).limit(5).all()

    # ... (существующий код до income_total)

    income_total = db.session.query(func.sum(IncomeTransaction.service_cost)).filter(
        IncomeTransaction.datetime >= start_datetime,
        IncomeTransaction.datetime < end_datetime
    ).scalar() or 0

    # --- РАСХОДЫ за выбранный месяц ---
    total_expenses = db.session.query(func.sum(Expense.amount)).filter(
        Expense.date >= start_date,
        Expense.date < end_date
    ).scalar() or 0

    # --- Чистая прибыль ---
    net_profit = income_total - total_expenses

    # --- Расходы по месяцам текущего года (для графика) ---
    monthly_expenses = []
    for m in range(1, 13):
        start_of_month = date(year, m, 1)
        if m == 12:
            end_of_month = date(year + 1, 1, 1)
        else:
            end_of_month = date(year, m + 1, 1)
        exp_m = db.session.query(func.sum(Expense.amount)).filter(
            Expense.date >= start_of_month,
            Expense.date < end_of_month
        ).scalar() or 0
        monthly_expenses.append(round(exp_m, 2))

    # --- Последние 5 расходов (вместо последних смен) ---
    recent_expenses = Expense.query.order_by(Expense.date.desc()).limit(5).all()

    income_total = db.session.query(func.sum(IncomeTransaction.service_cost)).filter(
        IncomeTransaction.datetime >= start_datetime,
        IncomeTransaction.datetime < end_datetime
    ).scalar() or 0

    # --- Расходы (операционные) за выбранный месяц ---
    total_expenses = db.session.query(func.sum(Expense.amount)).filter(
        Expense.date >= start_date,
        Expense.date < end_date
    ).scalar() or 0

    # --- Затраты на смены (зарплаты) за месяц ---
    #    (используем уже полученный shifts_month)
    total_salary_cost = sum(
        (s.hours or 0) * (s.rate or 0) + (s.bonus or 0) - (s.deduction or 0)
        for s in shifts_month
    )

    # --- Налог 8% от дохода ---
    tax_amount = income_total * 0.08

    # --- Чистая прибыль (доход - расходы - зарплаты - налог) ---
    net_profit = income_total - total_expenses - total_salary_cost - tax_amount

    # --- Расходы по месяцам текущего года (для графика) ---
    monthly_expenses = []
    for m in range(1, 13):
        start_of_month = date(year, m, 1)
        if m == 12:
            end_of_month = date(year + 1, 1, 1)
        else:
            end_of_month = date(year, m + 1, 1)
        exp_m = db.session.query(func.sum(Expense.amount)).filter(
            Expense.date >= start_of_month,
            Expense.date < end_of_month
        ).scalar() or 0
        monthly_expenses.append(round(exp_m, 2))

    # --- Последние расходы ---
    recent_expenses = Expense.query.order_by(Expense.date.desc()).limit(5).all()

    return render_template('dashboard.html',
                           year=year, month=month,
                           prev_year=prev_year, prev_month=prev_month,
                           next_year=next_year, next_month=next_month,
                           month_name=cal.month_name[month],
                           total_shifts=total_shifts,
                           total_hours=round(total_hours, 1),
                           income_total=round(income_total, 2),
                           total_expenses=round(total_expenses, 2),
                           total_salary_cost=round(total_salary_cost, 2),
                           tax_amount=round(tax_amount, 2),
                           net_profit=round(net_profit, 2),
                           months=months,
                           monthly_hours=monthly_hours,
                           monthly_income=monthly_income,
                           monthly_expenses=monthly_expenses,
                           top_employees=top_employees,
                           top_services=top_services,
                           recent_expenses=recent_expenses)


@app.route("/upload_wb", methods=["GET", "POST"])
@login_required
def upload_wb():
    if request.method == "POST":
        file = request.files.get("file")
        if file and allowed_income_file(file.filename):
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
            os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
            file.save(filepath)
            try:
                count = import_wildberries_excel(filepath)
                flash(f"Загружено {count} дней доходов Wildberries", "success")
            except Exception as e:
                flash(f"Ошибка: {str(e)}", "danger")
            return redirect(url_for("income_calendar"))
    return render_template("upload_wb.html")


@app.route("/api/income", methods=['GET'])
@login_required
def api_get_income():
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    source = request.args.get('source', 'all')  # all, yandex, wildberries
    if not year or not month:
        return jsonify({'error': 'Missing year/month'}), 400
    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1)
    else:
        end_date = datetime(year, month + 1, 1)

    query = IncomeTransaction.query.filter(
        IncomeTransaction.datetime >= start_date,
        IncomeTransaction.datetime < end_date
    )
    if source != 'all':
        query = query.filter(IncomeTransaction.source == source)

    transactions = query.all()
    daily = {}
    details_by_day = {}
    for t in transactions:
        date_str = t.datetime.date().isoformat()
        daily[date_str] = daily.get(date_str, 0) + (t.service_cost or 0)
        if date_str not in details_by_day:
            details_by_day[date_str] = []
        details_by_day[date_str].append({
            'service': t.service,
            'cost': t.service_cost,
            'order': t.order_ref,
            'source': t.source
        })
    result = [{'date': d, 'total': daily[d], 'details': details_by_day[d]} for d in daily]
    return jsonify(result)

@app.route("/employees", methods=["GET", "POST"])
@login_required
def employees_page():
    if request.method == "POST":
        file = request.files.get("file")
        if file and allowed_file(file.filename):
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
            os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
            file.save(filepath)
            try:
                import_excel(filepath)
                flash("Файл успешно загружен и обработан")
            except Exception as e:
                flash(f"Ошибка обработки файла: {str(e)}")
            return redirect(url_for("employees_page"))

    employees = Employee.query.all()
    # Для отображения смен используем сортировку по дате
    shifts = Shift.query.order_by(Shift.date.desc()).all()
    return render_template("employees.html", employees=employees, shifts=shifts)

@app.route("/employee/delete/<int:id>")
@login_required
def delete_employee(id):
    employee = Employee.query.get_or_404(id)
    # Удаляем все смены сотрудника (если не настроен каскад)
    Shift.query.filter_by(employee_id=id).delete()
    db.session.delete(employee)
    db.session.commit()
    flash(f"Сотрудник {employee.name} и его смены удалены", "danger")
    return redirect(url_for('employees_page'))

@app.route("/employee/add", methods=["POST"])
@login_required
def add_employee():
    name = request.form.get("name", "").strip()
    if not name:
        flash("Имя сотрудника не может быть пустым", "danger")
        return redirect(url_for("employees_page"))
    if Employee.query.filter_by(name=name).first():
        flash(f"Сотрудник с именем '{name}' уже существует", "warning")
        return redirect(url_for("employees_page"))
    employee = Employee(name=name)
    db.session.add(employee)
    db.session.commit()
    flash(f"Сотрудник '{name}' успешно добавлен", "success")
    return redirect(url_for("employees_page"))

@app.route("/employee/edit/<int:id>", methods=["GET", "POST"])
@login_required
def edit_employee(id):
    employee = Employee.query.get_or_404(id)
    if request.method == "POST":
        new_name = request.form.get("name", "").strip()
        if not new_name:
            flash("Имя сотрудника не может быть пустым", "danger")
            return redirect(url_for("employees_page"))
        existing = Employee.query.filter(Employee.name == new_name, Employee.id != id).first()
        if existing:
            flash(f"Сотрудник с именем '{new_name}' уже существует", "warning")
            return redirect(url_for("employees_page"))
        employee.name = new_name
        db.session.commit()
        flash(f"Сотрудник переименован в '{new_name}'", "success")
        return redirect(url_for("employees_page"))
    return render_template("edit_employee.html", employee=employee)

@app.route("/upload_income", methods=["GET", "POST"])
@login_required
def upload_income():
    if request.method == "POST":
        file = request.files.get("file")
        if file and allowed_income_file(file.filename):
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
            os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
            file.save(filepath)
            try:
                count = import_income_excel(filepath)
                flash(f"Загружено {count} транзакций доходов", "success")
            except Exception as e:
                flash(f"Ошибка: {str(e)}", "danger")
            return redirect(url_for("income_calendar"))
    return render_template("upload_income.html")

# Экспорт смен в Excel
@app.route("/export_shifts")
@login_required
def export_shifts():
    shifts = Shift.query.order_by(Shift.date.desc()).all()
    data = []
    for s in shifts:
        data.append({
            "Сотрудник": s.employee.name,
            "Дата": s.date.strftime("%Y-%m-%d"),
            "Начало": s.start_time or "",
            "Конец": s.end_time or "",
            "Часы": s.hours or 0,
            "Ставка": s.rate or 0,
            "Бонус": s.bonus or 0,
            "Вычет": s.deduction or 0,
            "Статус": "Присутствовал" if s.status else "Отсутствовал"
        })
    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Смены")
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="shifts.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Акт выполненных работ (зарплатная ведомость)
# ------------------- Акт выполненных работ -------------------
@app.route("/employee/<int:id>/salary_act")
@login_required
def salary_act(id):
    employee = Employee.query.get_or_404(id)
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    today = datetime.today()
    if not year:
        year = today.year
    if not month:
        month = today.month

    # Границы месяца
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year+1, 1, 1)
    else:
        end_date = date(year, month+1, 1)

    shifts = Shift.query.filter(Shift.employee_id == id, Shift.date >= start_date, Shift.date < end_date).all()
    total_hours = sum(s.hours or 0 for s in shifts)
    total_rate_amount = sum((s.hours or 0) * (s.rate or 0) for s in shifts)
    total_bonus = sum(s.bonus or 0 for s in shifts)
    total_deduction = sum(s.deduction or 0 for s in shifts)
    net = total_rate_amount + total_bonus - total_deduction

    # Навигация по месяцам
    if month == 1:
        prev_year, prev_month = year - 1, 12
    else:
        prev_year, prev_month = year, month - 1
    if month == 12:
        next_year, next_month = year + 1, 1
    else:
        next_year, next_month = year, month + 1

    # Загрузка конфигурации заказчика (файл в корне проекта)
    import json
    try:
        with open('instance/customer_config.json', 'r', encoding='utf-8') as f:
            customer = json.load(f)
    except FileNotFoundError:
        customer = {
            "name": "ИП Иванов Иван Иванович",
            "inn": "123456789012",
            "address": "г. Москва, ул. Примерная, д. 1",
            "director_title": "Индивидуальный предприниматель",
            "signatory_name": "Иванов И.И."
        }

    act_date = datetime.today().strftime('%d.%m.%Y')
    act_number = f"{employee.id}-{year}{month:02d}"

    return render_template('salary_act.html',
                           employee=employee,
                           year=year, month=month,
                           prev_year=prev_year, prev_month=prev_month,
                           next_year=next_year, next_month=next_month,
                           month_name=cal.month_name[month],
                           prev_month_name=cal.month_name[prev_month],
                           next_month_name=cal.month_name[next_month],
                           shifts=shifts,
                           total_hours=total_hours,
                           total_rate_amount=round(total_rate_amount, 2),
                           total_bonus=round(total_bonus, 2),
                           total_deduction=round(total_deduction, 2),
                           net=round(net, 2),
                           customer=customer,
                           act_date=act_date,
                           act_number=act_number)

@app.route("/income_calendar")
@login_required
def income_calendar():
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    today = datetime.today()
    if not year:
        year = today.year
    if not month:
        month = today.month

    # навигация
    if month == 1:
        prev_year, prev_month = year - 1, 12
    else:
        prev_year, prev_month = year, month - 1
    if month == 12:
        next_year, next_month = year + 1, 1
    else:
        next_year, next_month = year, month + 1

    return render_template('income_calendar.html',
                           year=year, month=month,
                           prev_year=prev_year, prev_month=prev_month,
                           next_year=next_year, next_month=next_month,
                           month_name=cal.month_name[month],
                           prev_month_name=cal.month_name[prev_month],
                           next_month_name=cal.month_name[next_month])


@app.route("/expenses")
@login_required
def expenses_page():
    # Получаем год и месяц из запроса (для фильтрации)
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    today = datetime.today()
    if not year:
        year = today.year
    if not month:
        month = today.month

    # Навигация по месяцам
    if month == 1:
        prev_year, prev_month = year - 1, 12
    else:
        prev_year, prev_month = year, month - 1
    if month == 12:
        next_year, next_month = year + 1, 1
    else:
        next_year, next_month = year, month + 1

    # Фильтруем расходы за выбранный месяц
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year+1, 1, 1)
    else:
        end_date = date(year, month+1, 1)
    expenses = Expense.query.filter(Expense.date >= start_date, Expense.date < end_date).order_by(Expense.date.desc()).all()

    # Сумма расходов по категориям (для круговой диаграммы)
    category_totals = db.session.query(Expense.category, func.sum(Expense.amount)).filter(
        Expense.date >= start_date, Expense.date < end_date
    ).group_by(Expense.category).all()
    categories = [cat for cat, _ in category_totals]
    amounts = [amt for _, amt in category_totals]

    # Общая сумма за месяц
    total_expenses = sum(e.amount for e in expenses)

    # Получаем список категорий из модели
    expense_categories = Expense.get_categories()

    return render_template('expenses.html',
                           year=year, month=month,
                           prev_year=prev_year, prev_month=prev_month,
                           next_year=next_year, next_month=next_month,
                           month_name=cal.month_name[month],
                           expenses=expenses,
                           total_expenses=total_expenses,
                           category_totals=category_totals,
                           categories=categories,
                           amounts=amounts,
                           expense_categories=expense_categories)  # ← добавили

@app.route("/expenses/add", methods=["POST"])
@login_required
def add_expense():
    category = request.form.get('category')
    amount = float(request.form.get('amount'))
    date_str = request.form.get('date')
    description = request.form.get('description', '')
    expense_date = datetime.strptime(date_str, '%Y-%m-%d').date()
    expense = Expense(category=category, amount=amount, date=expense_date, description=description)
    db.session.add(expense)
    db.session.commit()
    flash(f'Расход "{category}" на сумму {amount} руб. добавлен', 'success')
    # Возвращаемся на ту же страницу с тем же месяцем/годом
    return redirect(url_for('expenses_page', year=expense_date.year, month=expense_date.month))

@app.route("/expenses/delete/<int:id>")
@login_required
def delete_expense(id):
    expense = Expense.query.get_or_404(id)
    year, month = expense.date.year, expense.date.month
    db.session.delete(expense)
    db.session.commit()
    flash('Расход удалён', 'danger')
    return redirect(url_for('expenses_page', year=year, month=month))


# ------------------- Календарь и API -------------------
@app.route("/calendar")
@login_required
def calendar_view():
    # Получаем год и месяц из запроса, иначе текущие
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    today = datetime.today()
    if not year:
        year = today.year
    if not month:
        month = today.month

    # Вычисляем предыдущий и следующий месяцы для навигации
    if month == 1:
        prev_year = year - 1
        prev_month = 12
    else:
        prev_year = year
        prev_month = month - 1

    if month == 12:
        next_year = year + 1
        next_month = 1
    else:
        next_year = year
        next_month = month + 1

    employees = [{'id': e.id, 'name': e.name} for e in Employee.query.all()]
    return render_template('calendar.html',
                           year=year,
                           month=month,
                           prev_year=prev_year,
                           prev_month=prev_month,
                           next_year=next_year,
                           next_month=next_month,
                           month_name=cal.month_name[month],
                           prev_month_name=cal.month_name[prev_month],
                           next_month_name=cal.month_name[next_month],
                           employees=employees)



@app.route("/api/shifts")
@login_required
def api_get_shifts():
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    query = Shift.query
    if year and month:
        start_date = date(year, month, 1)
        if month == 12:
            end_date = date(year+1, 1, 1)
        else:
            end_date = date(year, month+1, 1)
        query = query.filter(Shift.date >= start_date, Shift.date < end_date)
    shifts = query.order_by(Shift.date.desc()).all()
    shifts_list = [{
        'id': s.id,
        'employeeName': s.employee.name,
        'employee': s.employee.name,
        'date': s.date.isoformat(),
        'hours': s.hours or 0,
        'rate': s.rate or 0,
        'bonus': s.bonus or 0,
        'deduction': s.deduction or 0,
        'status': s.status
    } for s in shifts]
    return jsonify(shifts_list)

@app.route("/api/shifts", methods=['POST'])
@login_required
def api_save_shift():
    """Создаёт или обновляет смену."""
    data = request.get_json()
    shift_id = data.get('id')
    employee_id = data.get('employee_id')
    shift_date = datetime.strptime(data.get('date'), '%Y-%m-%d').date()
    start_time = data.get('start_time')
    end_time = data.get('end_time')
    hours = float(data.get('hours', 0) or 0)
    rate = float(data.get('rate', 0) or 0)
    bonus = float(data.get('bonus', 0) or 0)
    deduction = float(data.get('deduction', 0) or 0)
    status = data.get('status', True)

    if shift_id:
        shift = Shift.query.get(shift_id)
        if shift:
            shift.employee_id = employee_id
            shift.date = shift_date
            shift.start_time = start_time
            shift.end_time = end_time
            shift.hours = hours
            shift.rate = rate
            shift.bonus = bonus
            shift.deduction = deduction
            shift.status = status
    else:
        shift = Shift(employee_id=employee_id, date=shift_date,
                      start_time=start_time, end_time=end_time,
                      hours=hours, rate=rate, bonus=bonus,
                      deduction=deduction, status=status)
        db.session.add(shift)
    db.session.commit()
    return jsonify({'success': True, 'id': shift.id})


@app.route("/api/shifts/<int:shift_id>", methods=['DELETE'])
@login_required
def api_delete_shift(shift_id):
    shift = Shift.query.get(shift_id)
    if shift:
        db.session.delete(shift)
        db.session.commit()
        return jsonify({'success': True})
    return jsonify({'error': 'Shift not found'}), 404


@app.route("/api/dashboard")
@login_required
def api_dashboard():
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    today = datetime.today()
    if not year:
        year = today.year
    if not month:
        month = today.month

    # Данные за выбранный месяц (KPI, топы, последние расходы)
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year + 1, 1, 1)
    else:
        end_date = date(year, month + 1, 1)

    shifts_month = Shift.query.filter(Shift.date >= start_date, Shift.date < end_date).all()
    total_hours = sum((s.hours or 0) for s in shifts_month)

    start_datetime = datetime(year, month, 1)
    end_datetime = datetime(year, month + 1, 1) if month < 12 else datetime(year + 1, 1, 1)
    income_month = db.session.query(func.sum(IncomeTransaction.service_cost)).filter(
        IncomeTransaction.datetime >= start_datetime,
        IncomeTransaction.datetime < end_datetime
    ).scalar() or 0

    expenses_month = db.session.query(func.sum(Expense.amount)).filter(
        Expense.date >= start_date,
        Expense.date < end_date
    ).scalar() or 0

    total_salary_cost = sum(
        (s.hours or 0) * (s.rate or 0) + (s.bonus or 0) - (s.deduction or 0)
        for s in shifts_month
    )

    # --- Налог 8% от дохода ---
    tax_amount = income_month * 0.08

    # --- Чистая прибыль (доход - расходы - зарплаты - налог) ---
    net_profit = income_month - expenses_month - total_salary_cost - tax_amount

    # Топ сотрудников за месяц
    emp_hours = {}
    for s in shifts_month:
        emp_hours[s.employee.name] = emp_hours.get(s.employee.name, 0) + (s.hours or 0)
    top_employees = sorted(emp_hours.items(), key=lambda x: x[1], reverse=True)[:5]
    top_employees_json = [{'name': name, 'value': hours} for name, hours in top_employees]

    # Топ услуг за месяц
    service_income = db.session.query(
        IncomeTransaction.service,
        func.sum(IncomeTransaction.service_cost)
    ).filter(
        IncomeTransaction.datetime >= start_datetime,
        IncomeTransaction.datetime < end_datetime
    ).group_by(IncomeTransaction.service).all()
    top_services = sorted(service_income, key=lambda x: x[1], reverse=True)[:5]
    top_services_json = [{'name': service, 'value': amount} for service, amount in top_services]

    # Последние расходы (5 записей)
    recent_expenses = Expense.query.order_by(Expense.date.desc()).limit(5).all()
    recent_expenses_json = [{
        'id': e.id,
        'date': e.date.isoformat(),
        'category': e.category,
        'amount': e.amount,
        'description': e.description
    } for e in recent_expenses]

    # --- Годовые данные ---
    months_names = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']
    monthly_hours = []
    monthly_income = []
    monthly_expenses = []
    for m in range(1, 13):
        start_m = date(year, m, 1)
        if m == 12:
            end_m = date(year + 1, 1, 1)
        else:
            end_m = date(year, m + 1, 1)
        # часы
        shifts_m = Shift.query.filter(Shift.date >= start_m, Shift.date < end_m).all()
        hours_m = sum((s.hours or 0) for s in shifts_m)
        monthly_hours.append(hours_m)
        # доход
        start_dt = datetime(year, m, 1)
        end_dt = datetime(year, m + 1, 1) if m < 12 else datetime(year + 1, 1, 1)
        inc_m = db.session.query(func.sum(IncomeTransaction.service_cost)).filter(
            IncomeTransaction.datetime >= start_dt,
            IncomeTransaction.datetime < end_dt
        ).scalar() or 0
        monthly_income.append(inc_m)
        # расходы
        exp_m = db.session.query(func.sum(Expense.amount)).filter(
            Expense.date >= start_m,
            Expense.date < end_m
        ).scalar() or 0
        monthly_expenses.append(exp_m)

    return jsonify({
        'totalHours': total_hours,
        'totalIncome': income_month,
        'totalExpenses': expenses_month,
        'netProfit': net_profit,
        'topEmployees': top_employees_json,
        'topServices': top_services_json,
        'recentExpenses': recent_expenses_json,
        'months': months_names,
        'monthlyHours': monthly_hours,
        'monthlyIncome': monthly_income,
        'monthlyExpenses': monthly_expenses
    })

@app.route("/api/expenses")
@login_required
def api_expenses():
    year = request.args.get('year', type=int)
    month = request.args.get('month', type=int)
    today = datetime.today()
    if not year:
        year = today.year
    if not month:
        month = today.month
    start_date = date(year, month, 1)
    if month == 12:
        end_date = date(year+1, 1, 1)
    else:
        end_date = date(year, month+1, 1)
    expenses = Expense.query.filter(Expense.date >= start_date, Expense.date < end_date).order_by(Expense.date.desc()).all()
    total_expenses = sum(e.amount for e in expenses)
    category_totals = db.session.query(Expense.category, func.sum(Expense.amount)).filter(
        Expense.date >= start_date, Expense.date < end_date
    ).group_by(Expense.category).all()
    categories = [cat for cat, _ in category_totals]
    amounts = [amt for _, amt in category_totals]
    expenses_list = [{
        'id': e.id,
        'date': e.date.isoformat(),
        'category': e.category,
        'amount': e.amount,
        'description': e.description
    } for e in expenses]
    return jsonify({
        'totalExpenses': total_expenses,
        'categories': categories,
        'amounts': amounts,
        'expenses': expenses_list
    })

@app.route("/api/login", methods=["POST"])
def api_login():
    data = request.get_json()
    username = data.get("username")
    password = data.get("password")
    user = User.query.filter_by(username=username).first()
    if user and check_password_hash(user.password, password):
        login_user(user)
        return jsonify({"success": True, "message": "Logged in"})
    return jsonify({"success": False, "message": "Invalid credentials"}), 401

@app.route("/api/logout", methods=["POST"])
@login_required
def api_logout():
    logout_user()
    return jsonify({"success": True})

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5430, debug=True)