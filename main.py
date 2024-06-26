import openpyxl.utils.exceptions
import modules as mods
import os

print(r"""
 ______   __  __     ______   __     __  __     __        
/\  == \ /\ \_\ \   /\  ___\ /\ \   /\_\_\_\   /\ \       
\ \  _-/ \ \____ \  \ \  __\ \ \ \  \/_/\_\/_  \ \ \____  
 \ \_\    \/\_____\  \ \_\    \ \_\   /\_\/\_\  \ \_____\ 
  \/_/     \/_____/   \/_/     \/_/   \/_/\/_/   \/_____/ 
                                                                                                 
                                               
PyFiXL v0.3.0-beta
Simple Excel workbook processing with Python, making personal finance a breeze.
Created by R. Davis, Folfchie on Github.
      """)

while True:
    choice = input('\n>>> ').lower()
    try:
        match choice:
            case "quit":
                break   
            case "exit":
                break
                
            case "help":
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
                  
            case "proc_income":
                print('\nEnter exact filename with extension:')
                mods.process_income_workbook(input('>>> '))
                
            case "proc_mtd":
                print('\nEnter exact filenames with extensions:')
                mods.process_mtd_figures(income_wb_name=input('Income workbook >>> '),
                                         expenses_wb_name=input('Expenses workbook >>> '),
                                         mtd_wb_name=input('MTD workbook >>> '))
                                        
            case "proc_ytd":
                print('\nEnter paths and values:')
                mods.process_ytd_figures(mtd_dir=input("Path for month-to-date workbooks >>> "),
                                         ytd_dir=input("Path for year-to-date workbooks >>> "),
                                         year=input("Year you'd like to process (e.g. 2024) >>> "))
                                        
            case "cwd":
                print(os.getcwd())
            
            case "cd":
                os.chdir(input("Enter path >>> "))
                
            case "ls":
                print(os.listdir())

            case "debug":
                mods.open_test_window()
                
            case _:
                print(f"\nCommand '{choice}' not found")
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
