# PynanceXL v0.1

**PynanceXL** is a simple program created to process financial Excel workbooks.

## Prerequisites

- **Python 3** with **openpyxl**
- **Excel**, **LibreOffice Calc**, or another program to view/edit `.xlsx` workbooks.

## Installation

Download `PynanceXL1.0.zip` and extract it. You now have two choices.

### Option A: Use Python Interpreter
Simply run `main.py` using your desired Python interpreter. This is best if you just want to try out PynanceXL.

### Option B: Custom bash script
Included in the zip is a file named `pynance`. This is a custom bash script meant to be stored in `/usr/local/bin` on Debian distros. 

To properly configure this, copy `pynance` and the folder `PynanceXLv0.1` to the aforementioned directory. Ensure that the file `pynance` is **outside** of the Pynance folder.

Next, run the command `sudo chmod +x pynance`.

If done correctly, you can now use the bash command `pynance` to run the program.

## Usage
I created this program as I desired a basic, simple, automated
personal finance program that utilizes Excel files.

To run the program, run `main.py` using your favorite Python interpreter.

3 example workbooks are provided in the zip file.

- `income.xlsx`
| Enter your income figures here.
- `expenses.xlsx`
| Enter your expense figures here.
- `mtd_totals.xlsx`
| Leave empty. Month-to-date totals go here.

You may alter the filenames, make copies, and enter data in the cells as you see fit. Please 
refrain from altering any formatting, this may break the program.

## Bug reporting

Report any bugs and/or suggestions to [Folfchie](https://www.github.com/Folfchie) on Github.
 
