#! python3
# Specific script to take output of O365 auth log
# and output needed data with deduplication
# Author: Josh Smith

import csv
import tkinter as tk
from tkinter import filedialog
import os

def remove_dup(a: list) -> None:
    """Remove duplicate items from a given list. For this script, if the IP
    is different, it will not be considered a duplicate. It will not delete any
    duplicates that are due to failed logon attempts.

    :param a: List to iterate through
    """
    i = 0
    while i < len(a):
        j = i + 1
        while j < len(a):
            if a[i] == a[j] and 'Failure' not in a[i]:
                del a[j]
            else:
                j += 1
        i += 1


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
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    del headers[28:]
    del headers[25]
    del headers[22]
    del headers[6:19]
    del headers[0:4]
    for row in reader:
        del row[28:]
        del row[25]
        del row[22]
        del row[6:19]
        del row[0:4]
        new_list.append(row)

remove_dup(new_list)

# Ask the user to select a single file name for saving.
# Can choose same and overwrite
# Commented this out due to user input. Rather select and go.
# However, this overwrites origin file as is

# out_file = filedialog.asksaveasfilename(parent=application_window,
#                                       initialdir=os.getcwd(),
#                                       title="Please select a file to save:",
#                                       filetypes=my_filetypes)

# open input file and write output (overwrite)
with open(in_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(headers)
    writer.writerows(new_list)

