from tkinter import *
from tkinter import ttk
import modules as mods


def open_test_window():
    root = Tk()
    frm = ttk.Frame(root, padding=10)
    frm.grid()
    ttk.Label(frm, text="Path to file/directory:").grid(column=1, row=0)
    ttk.Button(frm, text="Process Income Workbook",
               command=lambda: mods.process_income_workbook('placeholder.xlsx')).grid(column=1, row=1)
    ttk.Button(frm, text="Process MTD Workbooks",
               command=lambda: mods.process_mtd_figures('income_wb_name',
                                                        'expenses_wb_name',
                                                        'mtd_wb_name')).grid(column=1, row=2)
    ttk.Button(frm, text="Process YTD Workbook",
               command=lambda: mods.process_ytd_figures('mtd_dir',
                                                        'ytd_dir',
                                                        'year')).grid(column=1, row=3)
    ttk.Button(frm, text="Quit", command=root.destroy).grid(column=1, row=4)

    root.resizable(width=False, height=False)
    root.title('PyFiXL')
    root.geometry("225x400")
    root.mainloop()


if True:
    open_test_window()
