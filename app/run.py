from app import create_app, db
from app.models import User, Announcement, Payment, Expenditure, BalanceSheet, BalanceSheetItem
from werkzeug.security import generate_password_hash
import os
from datetime import date

app = create_app()

def migrate_database():
    """Add any new columns to existing tables that may not exist yet."""
    with app.app_context():
        import sqlalchemy as sa
        inspector = sa.inspect(db.engine)
        # Table may not exist yet on first run; db.create_all() in seed_database handles it
        if not inspector.has_table('expenditures'):
            return
        columns = [col['name'] for col in inspector.get_columns('expenditures')]
        if 'payment_type' not in columns:
            with db.engine.connect() as conn:
                conn.execute(sa.text("ALTER TABLE expenditures ADD COLUMN payment_type VARCHAR(20) DEFAULT 'Bank'"))
                conn.commit()
            print("Migration: added payment_type column to expenditures table.")


def seed_database():
    with app.app_context():
        db.create_all()

        if User.query.count() > 0:
            return

        admin = User(
            username='admin',
            password_hash=generate_password_hash('admin123'),
            role='admin',
            apartment_number=None,
            is_first_login=False
        )
        db.session.add(admin)

        for i in range(1, 31):
            apt_num = f'A-{100+i}'
            owner = User(
                username=apt_num.lower(),
                password_hash=generate_password_hash('Welcome'),
                role='owner',
                apartment_number=apt_num,
                is_first_login=True
            )
            db.session.add(owner)

        db.session.flush()

        admin_user = User.query.filter_by(username='admin').first()

        announcements = [
            Announcement(title='Welcome to Apartment Manager', content='This is the new apartment management portal. Please log in with your apartment number and default password "Welcome".', created_by_id=admin_user.id),
            Announcement(title='Maintenance Work - March 2025', content='There will be maintenance work on the water supply on March 15, 2025 from 10 AM to 2 PM.', created_by_id=admin_user.id),
            Announcement(title='Monthly Maintenance Fee', content='Please pay your monthly maintenance fee of Rs. 2000 by the 10th of each month.', created_by_id=admin_user.id),
        ]
        for ann in announcements:
            db.session.add(ann)

        for i in range(1, 11):
            apt_num = f'A-{100+i}'
            p = Payment(
                apartment_number=apt_num,
                amount=2000.0,
                payment_date=date(2025, 2, 10),
                month=2,
                year=2025,
                notes='Monthly maintenance fee'
            )
            db.session.add(p)

        expenditures_data = [
            Expenditure(description='Security Guard Salary', amount=15000, category='Salary', expenditure_date=date(2025, 2, 28), month=2, year=2025),
            Expenditure(description='Electricity Bill - Common Area', amount=3500, category='Utilities', expenditure_date=date(2025, 2, 25), month=2, year=2025),
            Expenditure(description='Lift Maintenance', amount=2000, category='Maintenance', expenditure_date=date(2025, 2, 20), month=2, year=2025),
            Expenditure(description='Cleaning Supplies', amount=800, category='Supplies', expenditure_date=date(2025, 2, 15), month=2, year=2025),
        ]
        for exp in expenditures_data:
            db.session.add(exp)

        bs = BalanceSheet(
            title='Monthly Balance Sheet - February 2025',
            month=2,
            year=2025,
            opening_balance=25000,
            notes='Regular monthly balance sheet',
            created_by_id=admin_user.id
        )
        db.session.add(bs)
        db.session.flush()

        bs_items = [
            BalanceSheetItem(balance_sheet_id=bs.id, item_type='income', description='Maintenance Fee (10 apartments)', amount=20000),
            BalanceSheetItem(balance_sheet_id=bs.id, item_type='income', description='Parking Fee', amount=3000),
            BalanceSheetItem(balance_sheet_id=bs.id, item_type='expense', description='Security Guard Salary', amount=15000),
            BalanceSheetItem(balance_sheet_id=bs.id, item_type='expense', description='Electricity Bill', amount=3500),
            BalanceSheetItem(balance_sheet_id=bs.id, item_type='expense', description='Lift Maintenance', amount=2000),
            BalanceSheetItem(balance_sheet_id=bs.id, item_type='expense', description='Cleaning Supplies', amount=800),
        ]
        for item in bs_items:
            db.session.add(item)

        income_total = sum(i.amount for i in bs_items if i.item_type == 'income')
        expense_total = sum(i.amount for i in bs_items if i.item_type == 'expense')
        bs.income_total = income_total
        bs.expense_total = expense_total
        bs.closing_balance = bs.opening_balance + income_total - expense_total

        db.session.commit()
        print("Database seeded successfully!")

if __name__ == '__main__':
    migrate_database()
    seed_database()
    app.run(debug=os.environ.get('FLASK_DEBUG', 'false').lower() == 'true', host='0.0.0.0', port=5000)
