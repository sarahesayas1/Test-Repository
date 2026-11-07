
import random
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Configuration
NUM_ACCOUNTS = 100
OUTPUT_FILE = "sample_financial_accounts.xlsx"

# Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Financial Accounts"

# Column headers
headers = [
    "Account ID",
    "Account Holder",
    "Account Type",
    "Starting Balance",
    "Monthly Income",
    "Monthly Expenses",
    "Ending Balance",
    "Account Status",
    "Last Transaction Date"
]

ws.append(headers)

# Sample data
first_names = ["Sarah", "John", "Emma", "Michael", "Olivia",
               "James", "Sophia", "Daniel", "Emily", "David"]

last_names = ["Smith", "Johnson", "Williams", "Brown",
              "Jones", "Garcia", "Miller", "Davis"]

account_types = [
    "Checking", "Savings", "Credit Card", "Investment"
]

statuses = ["Active", "Inactive", "Frozen"]

# Generate random financial accounts
for i in range(1, NUM_ACCOUNTS + 1):

    account_id = f"ACC-{i:05d}"

    name = (
        random.choice(first_names)
        + " "
        + random.choice(last_names)
    )

    account_type = random.choice(account_types)

    starting_balance = round(random.uniform(500, 50000), 2)
    monthly_income = round(random.uniform(1000, 12000), 2)
    monthly_expenses = round(random.uniform(500, 8000), 2)

    status = random.choices(
        statuses,
        weights=[80, 15, 5]
    )[0]

    last_transaction = (
        datetime.now()
        - timedelta(days=random.randint(0, 365))
    ).date()

    row_number = i + 1

    ws.append([
        account_id,
        name,
        account_type,
        starting_balance,
        monthly_income,
        monthly_expenses,
        f"=D{row_number}+E{row_number}-F{row_number}",
        status,
        last_transaction
    ])

# Style headers
header_fill = PatternFill(
    start_color="17365D",
    end_color="17365D",
    fill_type="solid"
)

for cell in ws[1]:
    cell.font = Font(color="FFFFFF", bold=True)
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center")

# Format currency columns
for row in range(2, NUM_ACCOUNTS + 2):
    for col in range(4, 8):
        ws.cell(row, col).number_format = '$#,##0.00'

    ws.cell(row, 9).number_format = "yyyy-mm-dd"

# Adjust column widths
widths = [16, 24, 18, 20, 20, 20, 20, 18, 24]

for col_num, width in enumerate(widths, start=1):
    ws.column_dimensions[
        get_column_letter(col_num)
    ].width = width

# Enable filtering and freeze header
ws.auto_filter.ref = ws.dimensions
ws.freeze_panes = "A2"

# Save workbook
wb.save(OUTPUT_FILE)

print(f"Created {OUTPUT_FILE} with {NUM_ACCOUNTS} accounts!")
