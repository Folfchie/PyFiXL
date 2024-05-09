import openpyxl.utils.exceptions
import modules as mods
import os

print(r"""
    ____                                   _  __ __ 
   / __ \__  ______  ____ _____  ________ | |/ // / 
  / /_/ / / / / __ \/ __ `/ __ \/ ___/ _ \|   // /  
 / ____/ /_/ / / / / /_/ / / / / /__/  __/   |/ /___
/_/    \__, /_/ /_/\__,_/_/ /_/\___/\___/_/|_/_____/
      /____/                                        
                                               
PynanceXL v0.2.0-beta
Simple Excel workbook processing with Python, making personal finance a breeze.
Created by R. Davis, Folfchie on Github.
      """)

while True:
    choice = input('\n>>> ').lower()
    try:
        if choice == 'quit' or choice == 'exit':
            break
        elif choice == 'help':
            print("""
Command | Description | Usage
proc_income: Used for processing income workbooks. Usage: proc_income
proc_mtd: Used for processing month-to-date workbooks. Usage: proc_mtd
proc_ytd: Used for processing year-to-date workbooks. Usage: proc_ytd
cwd: Used to view the current working directory. Usage: cwd
cd: Used to change the current working directory. Usage: cd
ls: List files and paths in current working directory. Usage: ls
quit: Exits the program. Usage: quit or exit
help: Displays a list of commands and usages. Usage: help
                  """)
        elif choice == 'proc_income':
            print('\nEnter exact filename with extension:')
            mods.process_income_workbook(input('>>> '))
        elif choice == 'proc_mtd':
            print('\nEnter exact filenames with extensions:')
            mods.process_mtd_figures(income_wb_name=input('Income workbook >>> '),
                                     expenses_wb_name=input('Expenses workbook >>> '),
                                     mtd_wb_name=input('MTD workbook >>> '))
        elif choice == 'proc_ytd':
            print('\nEnter paths and values:')
            mods.process_ytd_figures(mtd_dir=input("Path for month-to-date workbooks >>> "),
                                     ytd_dir=input("Path for year-to-date workbooks >>> "),
                                     year=input("Year you'd like to process (e.g. 2024) >>> "))
        elif choice == 'cwd':
            print(os.getcwd())
        elif choice == 'cd':
            os.chdir(input('Enter path >>> '))
        elif choice == 'ls':
            print(os.listdir())
        else:
            print('\nCommand not found.')
    except openpyxl.utils.exceptions.InvalidFileException:
        print("\nOne or more invalid file formats. Try again.")
    except FileNotFoundError:
        print("\nOne or more files/directories do not exist. Try again.")
    except TypeError:
        print("\nInvalid data type found. Operation failed. Check your workbooks?")
    except NotADirectoryError:
        print("\nThe directory you entered does not exist. Try again.")
    except PermissionError:
        print("\nInsufficient permissions. Operation failed.")
