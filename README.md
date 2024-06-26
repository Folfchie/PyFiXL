# PyFiXL v0.3.0-beta

**PyFiXL** is a simple program created to process financial Excel workbooks.

## Prerequisites

- **Python 3** with **openpyxl**
- **Excel**, **LibreOffice Calc**, or another program to view/edit `.xlsx` workbooks.

## Installation

Download **PyFiXL** from the `main` branch on Github. You now have two choices.

### Option A: Use Python Interpreter
Simply run `main.py` using your desired Python interpreter. This is best if you just want to try out PyFiXL.

### Option B: Custom bash script (Debian)
**Note**: This method requires the use of `sudo` user privileges.

#### Step 1:
In the **PyFiXL** folder
 is a bash script named `pyfixl`. Copy this file to `/usr/local/bin`.

#### Step 2:
If needed, make the file executable by running the command `chmod +x pyfixl`. This will make the file executable. 
Now the `pyfixl` command is ready to use.

#### Step 3:
Run the command `pyfixl -i` and you will be prompted to enter the path of the **PyFiXL** folder
. 
This will be wherever you downloaded it to, such as `/home/user/Downloads/PyFiXL`.

#### Check if PyFiXL is installed
To check if **PyFiXL** is properly installed, run the command `pyfixl -r`.
If you receive the error that `/usr/local/bin/PyFiXL/main.py` does not exist, try **Step 3** again and ensure the path is entered correctly.

## Usage
I created this program as I desired a basic, simple, automated
personal finance program that utilizes Excel files.

### Files
Included with **PyFiXL** is a folder named `templates`.
There are four workbook templates for you to copy and use.

- `income.xlsx`
| Enter your income figures here.
- `expenses.xlsx`
| Enter your expense figures here.
- `jan-mtd.xlsx`
| Leave empty. Month-to-date totals go here.
- `20xx-ytd.xlsx`
| Leave empty. Data stored here is pulled from `mtd` worksheets.

Rename the workbooks to suit your organization tastes. 
Bear in mind that the `mtd` and `ytd` workbooks should follow
a particular formatting scheme in order for **PyFiXL** to read them.
Examples below:
- `2024-ytd.xlsx`
- `mar-mtd.xlsx`
- `jun-mtd.xlsx`

It is discouraged to alter workbook formatting,
as this may break program. Do so at your own risk.

### Commands

#### Bash Commands
- `pyfixl -h` 
| Displays a list of `pyfixl` bash commands.
- `pyfixl -r`
| If `pyfixl` is installed, runs the program.
- `pyfixl -i`
| Installs the program into `/usr/local/bin`

#### Python Commands
- `help`
| Prints of full list of `pyfixl` python commands.
- `cd`
| Change current working directory.
- `cwd`
| Print current working directory.
- `ls`
| Print files and directories in CWD

For a full list of commands, feel free to try out **PyFiXL**.
Then, run the `help` command.

## Updating PyFiXL
To update **PyFiXL**, simply download the latest version from Github. 
Then, if needed, refer to **Option B: Step 3** under **Installation**.

## Bug reporting
Report any bugs and/or suggestions to [Folfchie](https://www.github.com/Folfchie) on Github.
