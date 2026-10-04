"""
Simple Banking System - Mini Project
------------------------------------
Command-line banking application in Python.
All account data is stored permanently in an Excel file (banking_system.xlsx)
using the openpyxl library.

Install requirement:  pip install openpyxl
Run the program:      python banking_system.py
"""

import os
from openpyxl import Workbook, load_workbook

FILE_NAME = "banking_system.xlsx"
HEADERS = ["Account Number", "Name", "Mobile Number", "PIN", "Balance"]


# ---------------------------------------------------------------
# Excel file setup
# ---------------------------------------------------------------
def setup_excel_file():
    """Create banking_system.xlsx with sample accounts if it does not exist."""
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.title = "Accounts"
        ws.append(HEADERS)
        # PIN and Mobile are stored as text so leading zeros are not lost
        ws.append([1001, "Rahul Patil", "9876543210", "1234", 5000])
        ws.append([1002, "Priya Sharma", "9876543211", "5678", 3000])
        wb.save(FILE_NAME)


# ---------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------
def get_amount(message):
    """Ask for an amount and return it as a number. Returns None if invalid."""
    try:
        amount = float(input(message))
    except ValueError:
        print("Invalid amount! Please enter numbers only.")
        return None
    if amount <= 0:
        print("Invalid amount! Amount must be greater than zero.")
        return None
    return amount


def find_account(account_number):
    """
    Search for an account number in the Excel file.
    Returns (workbook, sheet, row_number) if found, otherwise None.
    """
    wb = load_workbook(FILE_NAME)
    ws = wb.active
    for row in range(2, ws.max_row + 1):          # row 1 is the header
        if ws.cell(row=row, column=1).value == account_number:
            return wb, ws, row
    return None


def verify_pin(ws, row, pin):
    """Return True if the entered PIN matches the stored PIN."""
    return str(ws.cell(row=row, column=4).value) == pin


def login():
    """
    Ask for account number and PIN.
    Returns (workbook, sheet, row_number) if login is successful, otherwise None.
    """
    try:
        account_number = int(input("Enter Account Number: "))
    except ValueError:
        print("Invalid account number! Please enter numbers only.")
        return None

    result = find_account(account_number)
    if result is None:
        print("Account not found!")
        return None

    wb, ws, row = result
    pin = input("Enter 4-digit PIN: ").strip()
    if not verify_pin(ws, row, pin):
        print("Incorrect PIN!")
        return None

    return wb, ws, row


def generate_account_number():
    """Generate a unique account number (last account number + 1)."""
    wb = load_workbook(FILE_NAME)
    ws = wb.active
    last_number = 1000
    for row in range(2, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if isinstance(value, int) and value > last_number:
            last_number = value
    return last_number + 1


# ---------------------------------------------------------------
# Main banking operations
# ---------------------------------------------------------------
def create_account():
    print("\n--- Create Account ---")

    name = input("Enter Account Holder Name: ").strip()
    if name == "" or not name.replace(" ", "").isalpha():
        print("Invalid name! Use letters and spaces only.")
        return

    mobile = input("Enter Mobile Number (10 digits): ").strip()
    if not (mobile.isdigit() and len(mobile) == 10):
        print("Invalid mobile number! It must be exactly 10 digits.")
        return

    pin = input("Set a 4-digit PIN: ").strip()
    if not (pin.isdigit() and len(pin) == 4):
        print("Invalid PIN! It must be exactly 4 digits.")
        return

    initial_deposit = get_amount("Enter Initial Deposit: ")
    if initial_deposit is None:
        return

    account_number = generate_account_number()

    wb = load_workbook(FILE_NAME)
    ws = wb.active
    ws.append([account_number, name, mobile, pin, initial_deposit])
    wb.save(FILE_NAME)

    print("\nAccount created successfully!")
    print("Your Account Number is:", account_number)
    print("Please remember your Account Number and PIN.")


def deposit_money():
    print("\n--- Deposit Money ---")
    result = login()
    if result is None:
        return
    wb, ws, row = result

    amount = get_amount("Enter amount to deposit: ")
    if amount is None:
        return

    new_balance = ws.cell(row=row, column=5).value + amount
    ws.cell(row=row, column=5).value = new_balance
    wb.save(FILE_NAME)

    print("Amount deposited successfully!")
    print("Updated Balance: Rs.", new_balance)


def withdraw_money():
    print("\n--- Withdraw Money ---")
    result = login()
    if result is None:
        return
    wb, ws, row = result

    amount = get_amount("Enter amount to withdraw: ")
    if amount is None:
        return

    balance = ws.cell(row=row, column=5).value
    if amount > balance:
        print("Insufficient balance!")
        print("Available Balance: Rs.", balance)
        return

    new_balance = balance - amount
    ws.cell(row=row, column=5).value = new_balance
    wb.save(FILE_NAME)

    print("Amount withdrawn successfully!")
    print("Updated Balance: Rs.", new_balance)


def check_balance():
    print("\n--- Check Balance ---")
    result = login()
    if result is None:
        return
    wb, ws, row = result

    print("\nAccount Number :", ws.cell(row=row, column=1).value)
    print("Account Holder :", ws.cell(row=row, column=2).value)
    print("Current Balance: Rs.", ws.cell(row=row, column=5).value)


# ---------------------------------------------------------------
# Main program
# ---------------------------------------------------------------
def main():
    setup_excel_file()

    while True:
        print("\n===== SIMPLE BANKING SYSTEM =====")
        print("1. Create Account")
        print("2. Deposit Money")
        print("3. Withdraw Money")
        print("4. Check Balance")
        print("5. Exit")

        choice = input("Enter your choice (1-5): ").strip()

        if choice == "1":
            create_account()
        elif choice == "2":
            deposit_money()
        elif choice == "3":
            withdraw_money()
        elif choice == "4":
            check_balance()
        elif choice == "5":
            print("\nThank you for using the Patil's Banking System. Goodbye!")
            break
        else:
            print("Invalid choice! Please enter a number from 1 to 5.")


if __name__ == "__main__":
    main()
