# PynanceXL v0.1.0-beta

**PynanceXL** is a simple program created to process financial Excel workbooks.

**Note:** This program is not affiliated with the `pynance` python library. 

## Prerequisites

- **Python 3** with **openpyxl**
- **Excel**, **LibreOffice Calc**, or another program to view/edit `.xlsx` workbooks.

## Installation

Download the PynanceXL repository from Github. You now have two choices.

### Option A: Use Python Interpreter
Simply run `main.py` using your desired Python interpreter. This is best if you just want to try out PynanceXL.

### Option B: Custom bash script (Linux)
In the **PynanceXL** repository, there is a file named `pynance`. This is a custom bash script meant to be stored in `/usr/local/bin` on Debian distros.  This will be renamed in a later release to avoid confusion and conflict with the `pynance` python library.

To properly configure this, move the `pynanceXL` folder to `/usr/local/bin`. Then, from within the `pynanceXL` folder, move the file `pynance`
to `/usr/local/bin`. If done correctly, you should have a `pynance` file and a `pynanceXL` folder in `/usr/local/bin`.

Next, run the command `sudo chmod +x pynance`.

Now you can now use the bash command `pynance` to run the program.

## Usage
I created this program as I desired a basic, simple, automated
personal finance program that utilizes Excel files.

3 example workbooks are provided in the zip file.

- `income.xlsx`
| Enter your income figures here.
- `expenses.xlsx`
| Enter your expense figures here.
- `mtd_totals.xlsx`
| Leave empty. Month-to-date totals go here.

These serve as templates, copy them and rename them something descriptive, such as `2024-04-income.xlsx`. 

Be sure to run `pynanceXL` in the same directory as the files you wish to process.

Please refrain from altering any formatting, this may break the program.

## Bug reporting

Report any bugs and/or suggestions to [Folfchie](https://www.github.com/Folfchie) on Github.
