import openpyxl.utils.exceptions
import modules as mods

print(r"""
    ____                                   _  __ __ 
   / __ \__  ______  ____ _____  ________ | |/ // / 
  / /_/ / / / / __ \/ __ `/ __ \/ ___/ _ \|   // /  
 / ____/ /_/ / / / / /_/ / / / / /__/  __/   |/ /___
/_/    \__, /_/ /_/\__,_/_/ /_/\___/\___/_/|_/_____/
      /____/                                        
                                               
PynanceXL v0.1
Simple Excel workbook processing with Python, making personal finance a breeze.
Created by R. Davis, Folfchie on Github.
      """)

while True:
    choice = input('\n>>> ').lower()
    try:
        if choice == 'quit':
            break
        elif choice == 'help':
            print("""
Command | Description | Usage
proc_income: Used for processing income workbooks. Usage: proc_income
proc_mtd: Used for processing month-to-date workbooks. Usage: proc_mtd
quit: Exits the program. Usage: quit
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
        else:
            print('\nCommand not found.')
    except openpyxl.utils.exceptions.InvalidFileException:
        print("\nOne or more invalid file formats. Try again.")
    except FileNotFoundError:
        print("\nOne or more files do not exist. Try again.")
    except TypeError:
        print("\nInvalid data type found. Operation failed. Check your workbooks?")
