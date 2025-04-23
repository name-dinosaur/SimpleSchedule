# SimpleSchedule
An automated shift scheduling tool built in Python. Assigns employees to 8 hour morning, evening, or night shifts for 24/7 coverage, using depth-first search with backtracking to ensure fairness and compliance with constraints.

---

## Features

- Reads employee availability from Excel
- Assigns 2 employees per shift, 3 shifts/day
- Ensures no double shifts or night → morning transitions
- Gives priority to under-scheduled employees
- Generates a color-coded Excel schedule

---

## Input Format

Place the Excel file `Employee_Availability.xlsx` in the project folder.

---

## Output

An Excel file named `Schedule.xlsx` is generated in the same folder. It shows:
- One row per employee
- One column per day
- Color-coded shift assignments:
  - Morning → light yellow
  - Evening → light blue
  - Night → light red
  - Unassigned → black

---

## How to Run

### Option 1: Python Script
1. Install requirements:
   ```bash
   pip install pandas openpyxl
