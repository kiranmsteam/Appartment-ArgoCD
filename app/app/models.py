from datetime import datetime
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from app import db, login_manager


class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='owner')  # 'admin' or 'owner'
    apartment_number = db.Column(db.String(20), nullable=True)
    is_first_login = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    payments = db.relationship('Payment', backref='user', lazy=True, foreign_keys='Payment.user_id')
    announcements = db.relationship('Announcement', backref='creator', lazy=True)
    balance_sheets = db.relationship('BalanceSheet', backref='creator', lazy=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.username}>'


class BalanceSheet(db.Model):
    __tablename__ = 'balance_sheets'

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    month = db.Column(db.Integer, nullable=False)
    year = db.Column(db.Integer, nullable=False)
    income_total = db.Column(db.Float, default=0.0)
    expense_total = db.Column(db.Float, default=0.0)
    opening_balance = db.Column(db.Float, default=0.0)
    closing_balance = db.Column(db.Float, default=0.0)
    notes = db.Column(db.Text, nullable=True)
    file_data = db.Column(db.LargeBinary, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)

    items = db.relationship('BalanceSheetItem', backref='balance_sheet', lazy=True, cascade='all, delete-orphan')

    @property
    def month_name(self):
        from calendar import month_name
        return month_name[self.month]

    def __repr__(self):
        return f'<BalanceSheet {self.title}>'


class BalanceSheetItem(db.Model):
    __tablename__ = 'balance_sheet_items'

    id = db.Column(db.Integer, primary_key=True)
    balance_sheet_id = db.Column(db.Integer, db.ForeignKey('balance_sheets.id'), nullable=False)
    item_type = db.Column(db.String(10), nullable=False)  # 'income' or 'expense'
    description = db.Column(db.String(200), nullable=False)
    amount = db.Column(db.Float, nullable=False)

    def __repr__(self):
        return f'<BalanceSheetItem {self.item_type}: {self.description}>'


class Payment(db.Model):
    __tablename__ = 'payments'

    id = db.Column(db.Integer, primary_key=True)
    apartment_number = db.Column(db.String(20), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    amount = db.Column(db.Float, nullable=False)
    payment_date = db.Column(db.Date, nullable=False)
    month = db.Column(db.Integer, nullable=False)
    year = db.Column(db.Integer, nullable=False)
    notes = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<Payment {self.apartment_number} {self.amount}>'


class Expenditure(db.Model):
    __tablename__ = 'expenditures'

    id = db.Column(db.Integer, primary_key=True)
    description = db.Column(db.String(200), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    category = db.Column(db.String(100), nullable=True)
    payment_type = db.Column(db.String(20), nullable=True, default='Bank')  # 'Bank' or 'Cash'
    expenditure_date = db.Column(db.Date, nullable=False)
    month = db.Column(db.Integer, nullable=False)
    year = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<Expenditure {self.description}>'


class Announcement(db.Model):
    __tablename__ = 'announcements'

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    content = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)

    def __repr__(self):
        return f'<Announcement {self.title}>'


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))
