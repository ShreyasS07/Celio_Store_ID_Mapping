import pandas as pd
from datetime import datetime
from datetime import date
from datetime import timedelta
from tkinter import *
from tkinter import ttk
import time
from xlutils.copy import copy
from tkinter import ttk, StringVar, filedialog, messagebox
from tkinter.filedialog import askopenfilename

def select_raw_file():
    global raw_file
    raw = askopenfilename(filetypes=[('Excel Files', '*.xls;*.xlsx;*,csv;')])
    label_file_explorer.configure(text="File Opened: " + raw)
    raw_file = pd.read_excel(raw)
    raw_file.head()
    # return raw_file if bool(raw_file) else None
def select_count_file():
    global count
    count_file = askopenfilename(filetypes=[('Excel Files', '*.xls;*.xlsx;*,csv;')])
    label_file_explorer.configure(text="File Opened: " + count_file)
    count = pd.read_excel(count_file)
    count.head()
def celio_process():
    # New Data Frame with Required 4 columns
    count_df = count[['Site Name', 'Start date', 'Hour begin', 'Traffic']]
    # Changing the DATE format to YY-MM-DD
    dates = []
    count_df['Start date'] = pd.to_datetime(count_df['Start date']).dt.strftime("%y-%m-%d")
    for i in range(len(count_df['Start date'])):
        count_df['Start date'][i] = '20'+str(count_df['Start date'][i])
    for i, row in count_df.iterrows():
        start_date = pd.to_datetime(row['Start date'])
        yest = start_date - timedelta(1)
        yest = yest.date()
        dates.append(yest)
    print("Date created_on is",dates[0])
    # print("Printing length of 'Created_on' column",len(dates))
    def last_two_digits(raw_file, Code):
        return raw_file["Code"].apply(lambda x: int(str(x)[-2:]))
    # Add a new column with the last two digits of the values in the 'number' column
    raw_file['new_code'] = last_two_digits(raw_file, 'number')
    # Printing the raw DataFrame with new_code column
    raw_file.head()
    id_num = []
    for index, row in count_df.iterrows():
        if len(raw_file[raw_file['Site Name'] == row['Site Name']]) == 1:
            store_id = raw_file.index[raw_file['Site Name'] == row['Site Name']].values[0]
            id_num.append(raw_file.loc[store_id, "new_code"])
    # print("Length of Traffic column ", len(id_num))
    count_df['Hour begin'].head()
    click = count_df['Hour begin'].str.split(':')
    count_df['Hour begin'] = click
    time = []
    for i in range (len(count_df['Hour begin'])):
        #print(count_df['Hour begin'][i][0])
        t = (count_df['Hour begin'][i][0])
        time.append(t)
    # print("Length of Hour column ", len(time))
    visitor = []
    visitor = list(count_df['Traffic'])
    # print("Length of Traffic column ", len(visitor))
    Data = pd.DataFrame({'created_on': dates,
                         'hour': time,
                         'visitors': visitor,
                        'store_id': id_num}, columns=['created_on', 'store_id', 'hour', 'visitors'])
    with pd.ExcelWriter("Celio_store_id_mapping.xlsx") as writer:
        Data.to_excel(writer, index=False)
window = Tk()
window.title(" Mindful Automation Pvt Ltd ")
window.geometry('600x100')
window.configure(background="white")
label_file_explorer = Label(window, text=" Celio Store ID Mapping ",
                            width=100, height=3,
                            fg="blue")
Footfall_file = Button(window,
                     text=" Select the footfall Footfall Mapping file ",
                     command=select_raw_file)
Count_file = Button(window,
                     text=" Select the Daily Hourly Count file ",
                     command=select_count_file)
Start = Button(window,
                     text=" Submit ",
                     command=celio_process)
label_file_explorer.grid(column=1, row=1)
Footfall_file.grid(column=1,row=2)
Count_file.grid(column=1, row=3)
Start.grid(column=1, row=4)
window.mainloop()