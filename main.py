import datetime
import os
from fileinput import filename
from time import sleep
from win32_setctime import setctime
import tkinter as tk
import tkcalendar as tc
from tkinter import filedialog as fd
import openpyxl
import calendar
from babel.dates import format_date, parse_date, get_day_names, get_month_names
from babel.numbers import *

filenamee = ''


def change(filename, fileendname, author,
           time):
    fh = openpyxl.load_workbook(filename)
    obj = fh.properties  # To get old properties
    fh.properties.creator = author
    fh.properties.created = time
    fh.properties.modified = time
    fh.properties.lastModifiedBy = author
    print(fh.properties)
    fh.save(fileendname)
    sleep(3)
    os.utime(fileendname, (time.timestamp(), time.timestamp()))
    setctime(fileendname, time.timestamp())
    label.configure(text="Готово!")


def submit():
    global filenamee
    print(filenamee)
    if author_e.get() and fileendname_e.get() and time.get_date() and min_sb.get() and sec_hour.get() and sec.get() and filename != '':

        timecreation = []
        temp = ''
        for idx, i in enumerate(time.get_date()):
            if i != '/':
                temp += str(i)
            else:
                timecreation.append(int(temp))
                temp = ''
            if idx == len(time.get_date()) - 1:
                timecreation.append(int(temp))

        print(author_e.get())
        print(fileendname_e.get())
        print(time.get_date())

        change(filenamee, fileendname_e.get() + ".xlsx", author_e.get(),
               datetime.datetime(year=2000 + timecreation[2], month=timecreation[0], day=timecreation[1],
                                 hour=int(min_sb.get()), minute=int(sec_hour.get()), second=int(sec.get())))

    else:
        label.configure(text="Не введены значения, или введены не все!")


def select_file():
    global filenamee
    filetypes = (
        ('xlsx', '*.xlsx'),
    )

    file = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    filenamee = file

    open_button.configure(text=filenamee)


window = tk.Tk()
# window.geometry("300x250")
window.configure(background='white')
window.title("changer")
label = tk.Label(
    text="Изменить метаданные xlsx",
    height=3,
    width=50
)

inserttimetext = tk.Label(
    text="Время (Ч, М, С)",
    bg="white",
)

open_button = tk.Button(
    window,
    text='Open a File',
    command=select_file
)

fileendname = tk.Label(
    text="Имя выходного файла (без расширения)",
    bg="white",
)

fileendname_e = tk.Entry()

author = tk.Label(
    text="Автор",
    bg="white",
)

author_e = tk.Entry()

submit = tk.Button(
    text="Подтвердить!",
    width=16,
    height=1,
    command=submit
)

time = tc.Calendar()

fone = tk.Frame(window)
ftwo = tk.Frame(window)
hour_string = tk.StringVar()
min_string = tk.StringVar()
last_value_sec = ""
last_value = ""
f = ('Times', 15)

min_sb = tk.Spinbox(
    ftwo,
    from_=0,
    to=23,
    wrap=True,
    textvariable=hour_string,
    width=2,
    font=f,
    justify=tk.CENTER
)
sec_hour = tk.Spinbox(
    ftwo,
    from_=0,
    to=59,
    wrap=True,
    textvariable=min_string,
    font=f,
    width=2,
    justify=tk.CENTER
)

sec = tk.Spinbox(
    ftwo,
    from_=0,
    to=59,
    wrap=True,

    width=2,
    font=f,
    justify=tk.CENTER
)

label.pack()
open_button.pack(pady=4)

author.pack()
author_e.pack()
fileendname.pack()
fileendname_e.pack()

time.pack(pady=5)

min_sb.pack(side=tk.LEFT, fill=tk.X, expand=True)
sec_hour.pack(side=tk.LEFT, fill=tk.X, expand=True)
sec.pack(side=tk.LEFT, fill=tk.X, expand=True)

fone.pack(pady=10)
ftwo.pack(pady=10)

inserttimetext.pack()
submit.pack(pady=10)

window.mainloop()
