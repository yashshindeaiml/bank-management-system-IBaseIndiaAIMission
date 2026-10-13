# Bank Management System

A console-based Python mini-project that uses functions for banking operations and an Excel workbook for persistent storage.

## Setup

```powershell
python -m pip install -r requirements.txt
python bank_management.py
```

The first run creates `bank_data.xlsx` beside the Python program. It contains `Accounts` and `Transactions` worksheets. Accounts and successful deposits or withdrawals are saved immediately.

Each deposit or withdrawal is limited to Rs. 10,00,000.00. Run the regression tests with `python -m unittest -v`.

## Menu

1. Create an account
2. Deposit money
3. Withdraw money
4. View a balance
5. Search for an account
6. Display all accounts
7. View transaction history
8. Generate an account statement
9. Count accounts
10. Exit

This is an educational demonstration, not real banking software. It does not authenticate users or protect personal information; use fictional customer details only.