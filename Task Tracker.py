# Task Management Program
# The program calculates the hours spent on a task.
# It then calculates the amount of money made for each task at 1hour=$5.
# Time entered and information from the program are be stored in an excel file for future referencing.


from tkinter import *
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import pathlib

# create excel sheet or check if its already created
wb = pathlib.Path("Task_tracker_Data.xlsx")
if wb.exists():
    pass
else:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Date(Y-M-D)"
    ws["B1"] = "Start Time(H:M:S)"
    ws["C1"] = "End Time(H:M:S"
    ws["D1"] = "Hours Spent(Hrs)"
    ws["E1"] = "Amount Made($)"

    wb.save("Task_tracker_Data.xlsx")


class Tracker:

    def __init__(self, master):
        master.title("Task Tracker App")
        master.resizable(False, False)

        # Header Frame
        self.header_frame = ttk.Frame(master)
        self.header_frame.pack()

        ttk.Label(self.header_frame, text="Welcome to Task Tracker App").grid(row=0, column=0, columnspan=2)

        # Date Frame
        self.date_frame = ttk.Frame(master)
        self.date_frame.pack()

        # Declare label name for the Date frame fields
        ttk.Label(self.date_frame, text='Choose date').grid(row=0, column=0, padx=5, sticky='sw')

        # Declare input Date frame fields
        self.entry_date = DateEntry(self.date_frame, width=24, background='darkblue', foreground='white', borderwidth=2)

        # Display field Date frame widgets
        self.entry_date.grid(row=0, column=1, padx=5)

        # Time Frame
        self.time_frame = ttk.Frame(master)
        self.time_frame.pack()

        # Declare label name for the time frame fields
        ttk.Label(self.time_frame, text='Hrs(24hrs)').grid(row=0, column=1, padx=5, sticky='sw')
        ttk.Label(self.time_frame, text='Mins').grid(row=0, column=2, padx=5, sticky='sw')
        ttk.Label(self.time_frame, text='Secs').grid(row=0, column=3, padx=5, sticky='sw')
        ttk.Label(self.time_frame, text='Start Time:').grid(row=1, column=0, padx=5, sticky='sw')
        ttk.Label(self.time_frame, text='End Time:').grid(row=2, column=0, padx=5, sticky='sw')

        # Declare input Time frame fields
        self.hr_start = Spinbox(self.time_frame, from_=0, to_=23, width=8, state='readonly', justify=CENTER)
        self.min_start = Spinbox(self.time_frame, from_=0, to_=59, width=8, state='readonly', justify=CENTER)
        self.sec_start = Spinbox(self.time_frame, from_=0, to_=59, width=8, state='readonly', justify=CENTER)

        self.hr_end = Spinbox(self.time_frame, from_=0, to_=23, width=8, state='readonly', justify=CENTER)
        self.min_end = Spinbox(self.time_frame, from_=0, to_=59, width=8, state='readonly', justify=CENTER)
        self.sec_end = Spinbox(self.time_frame, from_=0, to_=59, width=8, state='readonly', justify=CENTER)

        # Display field Time frame widgets
        self.hr_start.grid(row=1, column=1, padx=5)
        self.min_start.grid(row=1, column=2, padx=5)
        self.sec_start.grid(row=1, column=3, padx=5)

        self.hr_end.grid(row=2, column=1, padx=5)
        self.min_end.grid(row=2, column=2, padx=5)
        self.sec_end.grid(row=2, column=3, padx=5)

        # Button Frame
        self.button_frame = ttk.Frame(master)
        self.button_frame.pack()

        # Declare buttons
        ttk.Button(self.button_frame, text='Submit Task', command=self.submit).grid(row=0, column=0, padx=5, sticky='e')
        ttk.Button(self.button_frame, text='Reset Task', command=self.reset).grid(row=0, column=1, padx=5, sticky='w')

    def submit(self):
        date = self.entry_date.get_date()
        date_str = date.strftime('%Y-%m-%d')

        hr_start = self.hr_start.get()
        min_start = self.min_start.get()
        sec_start = self.sec_start.get()

        hr_end = self.hr_end.get()
        min_end = self.min_end.get()
        sec_end = self.sec_end.get()

        s_time = hr_start + ":" + min_start + ":" + sec_start
        e_time = hr_end + ":" + min_end + ":" + sec_end

        start_time = date_str + " " + s_time
        end_time = date_str + " " + e_time

        starting_timestamp = datetime.strptime(start_time, '%Y-%m-%d %H:%M:%S')
        ending_timestamp = datetime.strptime(end_time, '%Y-%m-%d %H:%M:%S')

        time_spent = ending_timestamp - starting_timestamp

        hrs_spent = time_spent.total_seconds()/3600

        amount_made = hrs_spent*5

        print(hrs_spent)
        print(amount_made)

        wb = openpyxl.load_workbook("Task_tracker_Data.xlsx")
        ws = wb.active
        ws.cell(column=1, row=ws.max_row + 1, value=date_str)
        ws.cell(column=2, row=ws.max_row, value=s_time)
        ws.cell(column=3, row=ws.max_row, value=e_time)
        ws.cell(column=4, row=ws.max_row, value=hrs_spent)
        ws.cell(column=5, row=ws.max_row, value=amount_made)

        wb.save("Task_tracker_Data.xlsx")

    def reset(self):
        pass
        # self.entry_start_time.delete(0, 'end')
        # self.entry_end_time.delete(0, 'end')


def main():
    root = Tk()
    root.geometry("500x500")
    tracker = Tracker(root)
    root.mainloop()


if __name__ == "__main__":
    main()
