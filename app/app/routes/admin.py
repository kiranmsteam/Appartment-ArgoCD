from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required, current_user
from datetime import date, datetime
from app import db
from app.models import User, Payment, Expenditure, BalanceSheet, BalanceSheetItem, Announcement

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')

# Apartment numbers from the Reference Sheet (Cr side - credit/payment received)
APARTMENT_NUMBERS = [
    'KF1', 'KF2', 'KF3', 'KF4',
    'KG1', 'KG2', 'KG3', 'KG4',
    'KS1', 'KS2', 'KS3', 'KS4',
    'RF1', 'RF2', 'RF3',
    'RG1', 'RG2', 'RG3',
    'RS1', 'RS2', 'RS3',
    'VF1', 'VF2', 'VF3',
    'VG1', 'VG2', 'VG3',
    'VS1', 'VS2', 'VS3',
]

# Expense categories from the Reference Sheet (Dr side - debit/expense)
EXPENSE_CATEGORIES = [
    'Bank',
    'Bank Interest',
    'Bank FD',
    'BESCOM',
    'BWSSB',
    'Garden',
    'Generator',
    'Lift',
    'Maintenance',
    'Security',
    'Other',
]


def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role != 'admin':
            flash('Access denied. Admin privileges required.', 'danger')
            return redirect(url_for('auth.login'))
        return f(*args, **kwargs)
    return decorated


@admin_bp.route('/dashboard')
@login_required
@admin_required
def dashboard():
    today = date.today()
    current_month = today.month
    current_year = today.year

    total_payments = db.session.query(db.func.sum(Payment.amount)).filter_by(
        month=current_month, year=current_year
    ).scalar() or 0.0

    total_expenditure = db.session.query(db.func.sum(Expenditure.amount)).filter_by(
        month=current_month, year=current_year
    ).scalar() or 0.0

    balance = total_payments - total_expenditure

    apartments_paid = db.session.query(db.func.count(db.func.distinct(Payment.apartment_number))).filter_by(
        month=current_month, year=current_year
    ).scalar() or 0

    announcements_count = Announcement.query.count()

    recent_payments = Payment.query.order_by(Payment.created_at.desc()).limit(5).all()
    recent_expenditures = Expenditure.query.order_by(Expenditure.created_at.desc()).limit(5).all()

    return render_template('admin/dashboard.html',
                           total_payments=total_payments,
                           total_expenditure=total_expenditure,
                           balance=balance,
                           apartments_paid=apartments_paid,
                           announcements_count=announcements_count,
                           recent_payments=recent_payments,
                           recent_expenditures=recent_expenditures,
                           current_month=current_month,
                           current_year=current_year)


@admin_bp.route('/balance-sheet', methods=['GET', 'POST'])
@login_required
@admin_required
def balance_sheet():
    if request.method == 'POST':
        title = request.form.get('title', '').strip()
        month = int(request.form.get('month', 1))
        year = int(request.form.get('year', date.today().year))
        opening_balance = float(request.form.get('opening_balance', 0) or 0)
        notes = request.form.get('notes', '').strip()

        income_descriptions = request.form.getlist('income_description[]')
        income_amounts = request.form.getlist('income_amount[]')
        expense_descriptions = request.form.getlist('expense_description[]')
        expense_amounts = request.form.getlist('expense_amount[]')

        bs = BalanceSheet(
            title=title,
            month=month,
            year=year,
            opening_balance=opening_balance,
            notes=notes,
            created_by_id=current_user.id
        )
        db.session.add(bs)
        db.session.flush()

        income_total = 0.0
        for desc, amt in zip(income_descriptions, income_amounts):
            if desc.strip() and amt:
                amount = float(amt)
                income_total += amount
                item = BalanceSheetItem(
                    balance_sheet_id=bs.id,
                    item_type='income',
                    description=desc.strip(),
                    amount=amount
                )
                db.session.add(item)

        expense_total = 0.0
        for desc, amt in zip(expense_descriptions, expense_amounts):
            if desc.strip() and amt:
                amount = float(amt)
                expense_total += amount
                item = BalanceSheetItem(
                    balance_sheet_id=bs.id,
                    item_type='expense',
                    description=desc.strip(),
                    amount=amount
                )
                db.session.add(item)

        bs.income_total = income_total
        bs.expense_total = expense_total
        bs.closing_balance = opening_balance + income_total - expense_total

        db.session.commit()
        flash('Balance sheet created successfully!', 'success')
        return redirect(url_for('admin.view_balance_sheet', id=bs.id))

    return render_template('admin/balance_sheet.html', today_year=date.today().year)


@admin_bp.route('/balance-sheet/<int:id>')
@login_required
@admin_required
def view_balance_sheet(id):
    bs = BalanceSheet.query.get_or_404(id)
    income_items = [i for i in bs.items if i.item_type == 'income']
    expense_items = [i for i in bs.items if i.item_type == 'expense']
    return render_template('admin/view_balance_sheet.html', bs=bs,
                           income_items=income_items, expense_items=expense_items)


@admin_bp.route('/balance-sheets')
@login_required
@admin_required
def balance_sheets():
    sheets = BalanceSheet.query.order_by(BalanceSheet.year.desc(), BalanceSheet.month.desc()).all()
    return render_template('admin/balance_sheets.html', sheets=sheets)


@admin_bp.route('/payments', methods=['GET', 'POST'])
@login_required
@admin_required
def payments():
    if request.method == 'POST':
        apartment_number = request.form.get('apartment_number', '').strip()
        amount = float(request.form.get('amount', 0) or 0)
        payment_date_str = request.form.get('payment_date', '')
        month = int(request.form.get('month', 1))
        year = int(request.form.get('year', date.today().year))
        notes = request.form.get('notes', '').strip()

        payment_date = datetime.strptime(payment_date_str, '%Y-%m-%d').date() if payment_date_str else date.today()

        user = User.query.filter_by(apartment_number=apartment_number).first()
        payment = Payment(
            apartment_number=apartment_number,
            user_id=user.id if user else None,
            amount=amount,
            payment_date=payment_date,
            month=month,
            year=year,
            notes=notes
        )
        db.session.add(payment)
        db.session.commit()
        flash(f'Payment recorded for apartment {apartment_number}.', 'success')
        return redirect(url_for('admin.payments'))

    recent_payments = Payment.query.order_by(Payment.created_at.desc()).limit(10).all()
    return render_template('admin/payments.html', payments=recent_payments,
                           today=date.today(), today_year=date.today().year,
                           apartment_numbers=APARTMENT_NUMBERS)


@admin_bp.route('/expenditure', methods=['GET', 'POST'])
@login_required
@admin_required
def expenditure():
    if request.method == 'POST':
        description = request.form.get('description', '').strip()
        amount = float(request.form.get('amount', 0) or 0)
        category = request.form.get('category', '').strip()
        payment_type = request.form.get('payment_type', 'Bank').strip()
        expenditure_date_str = request.form.get('expenditure_date', '')
        month = int(request.form.get('month', 1))
        year = int(request.form.get('year', date.today().year))

        expenditure_date = datetime.strptime(expenditure_date_str, '%Y-%m-%d').date() if expenditure_date_str else date.today()

        exp = Expenditure(
            description=description,
            amount=amount,
            category=category,
            payment_type=payment_type,
            expenditure_date=expenditure_date,
            month=month,
            year=year
        )
        db.session.add(exp)
        db.session.commit()
        flash('Expenditure recorded successfully.', 'success')
        return redirect(url_for('admin.expenditure'))

    recent_expenditures = Expenditure.query.order_by(Expenditure.created_at.desc()).limit(10).all()
    return render_template('admin/expenditure.html', expenditures=recent_expenditures,
                           today=date.today(), today_year=date.today().year,
                           expense_categories=EXPENSE_CATEGORIES)


@admin_bp.route('/transactions')
@login_required
@admin_required
def transactions():
    all_payments = Payment.query.order_by(Payment.payment_date.desc()).all()
    all_expenditures = Expenditure.query.order_by(Expenditure.expenditure_date.desc()).all()
    return render_template('admin/transactions.html',
                           payments=all_payments, expenditures=all_expenditures)


@admin_bp.route('/announcements', methods=['GET', 'POST'])
@login_required
@admin_required
def announcements():
    if request.method == 'POST':
        title = request.form.get('title', '').strip()
        content = request.form.get('content', '').strip()
        if title and content:
            ann = Announcement(title=title, content=content, created_by_id=current_user.id)
            db.session.add(ann)
            db.session.commit()
            flash('Announcement posted successfully.', 'success')
        else:
            flash('Title and content are required.', 'danger')
        return redirect(url_for('admin.announcements'))

    all_announcements = Announcement.query.order_by(Announcement.created_at.desc()).all()
    return render_template('admin/announcements.html', announcements=all_announcements)


@admin_bp.route('/announcement/<int:id>/delete', methods=['POST'])
@login_required
@admin_required
def delete_announcement(id):
    ann = Announcement.query.get_or_404(id)
    db.session.delete(ann)
    db.session.commit()
    flash('Announcement deleted.', 'success')
    return redirect(url_for('admin.announcements'))
