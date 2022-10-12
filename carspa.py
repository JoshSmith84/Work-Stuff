#! python3
# Specific script to take output of O365 auth log and output needed data with deduplication
# Author: Josh Smith

import csv
import tkinter as tk
from tkinter import filedialog
import os
from remove_dup import remove_dup


application_window = tk.Tk()
new_list = []
my_filetypes = [('csv files', '.csv')]


# Ask the user to select a single file name.
in_file = filedialog.askopenfilename(parent=application_window,
                                    initialdir=os.getcwd(),
                                    title="Please select an input file:",
                                    filetypes=my_filetypes)


# open auth report csv and pull only what is needed
with open(in_file, encoding='utf-8', newline='') as csv_file:
    # headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        new_list.append(row)
        print(row)

remove_dup(new_list)
#open input file and write output (overwrite)
with open(in_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    # writer.writerow(headers)
    writer.writerows(new_list)

