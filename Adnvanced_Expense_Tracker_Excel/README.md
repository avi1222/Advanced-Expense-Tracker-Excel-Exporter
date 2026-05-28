# Expense Tracker with Excel Exporter

## Project Overview

This is a Python-based Expense Tracker application that allows users to:

- Add expenses
- Edit expenses
- Delete expenses
- Search expenses
- View expense summaries
- Export expenses to an Excel sheet
- Visualize expenses using charts
- Set category budgets

The project uses Python concepts such as:
- Lists
- Dictionaries
- Functions
- File Handling
- Exception Handling
- External Libraries

---

# Features

## 1. Add Expense
Users can add:
- Amount
- Description
- Date
- Category

Custom categories can also be added.

---

## 2. View Expenses
Displays all saved expenses with:
- Date
- Description
- Amount
- Category

---

## 3. Edit Expense
Allows modification of existing expenses.

---

## 4. Delete Expense
Allows deletion of expenses.

---

## 5. Search Expense
Search expenses using keywords.

---

## 6. Expense Summary
Shows:
- Total spending
- Category-wise spending
- Average expenses
- Highest expense
- Lowest expense

---

## 7. Export to Excel
Exports all expense records into:

expenses.xlsx

---

## 8. Expense Chart
Displays a pie chart of expenses by category using Matplotlib.

---

## 9. Budget Tracking
Users can:
- Set category budgets
- Receive warnings when budget is exceeded

---

# Technologies Used

- Python
- JSON
- openpyxl
- matplotlib

---

# Libraries Required

Install required libraries using:

```bash
pip install openpyxl matplotlib
```

---

# How to Run the Program

Open terminal or command prompt.

Run:

```bash
python expense_tracker.py
```

---

# Menu Options

When the program starts, the following menu appears:

```text
===== Expense Tracker =====

1. Add Expense
2. View Expenses
3. Edit Expense
4. Delete Expense
5. Search Expense
6. Expense Summary
7. Export To Excel
8. Show Expense Chart
9. Set Budget
10. Exit
```

---

# How to Use

## Add Expense
Choose option 1.

Enter:
- Amount
- Description
- Date
- Category

Example:

```text
Enter amount: ₹250
Enter description: Pizza
Enter date: 2025-05-25
Choose category: Entertainment
```

---

## View Expenses
Choose option 2.

Displays all saved expenses.

---

## Edit Expense
Choose option 3.

Select expense number and enter updated details.

---

## Delete Expense
Choose option 4.

Select expense number to delete.

---

## Search Expense
Choose option 5.

Enter keyword(description).

Example:

```text
pizza
```

---

## Expense Summary
Choose option 6.

Displays:
- Total spending
- Category-wise spending
- Average expense
- Highest expense
- Lowest expense

---

## Export to Excel
Choose option 7.

Creates:

```text
expenses.xlsx
```

---

## Show Expense Chart
Choose option 8.

Displays pie chart of category expenses.

---

## Set Budget
Choose option 9.

Enter:
- Category
- Budget amount

The system warns if expenses exceed budget.

---

# Data Structures Used

## List
Stores all expenses.

Example:

```python
expenses = []
```

---

## Dictionary
Stores each expense.

Example:

```python
{
    "amount": 250,
    "description": "Pizza",
    "date": "2025-05-25",
    "category": "Food"
}
```

---

# Files Generated

## expenses.json
Stores expense data permanently.

## expenses.xlsx
Excel export file.

---

# Sample Output

```text
Expense added successfully!

Expense Summary
--------------------------------

Total Spending: ₹1200

Food: ₹500
Transportation: ₹300
Entertainment: ₹400
```

---

# Author
Avi