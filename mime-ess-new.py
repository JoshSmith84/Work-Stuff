# mime-ess-new.py
# Author: Josh Smith
#
# take output sender list from mimecast, edit down to unanimous votes,
# and format for Barracuda
# In mimecast, output the list with all columns.
# Folder line should be edited for output
# as selecting output file is not coded yet


import csv
import tkinter as tk
from tkinter import filedialog
import os
import time


def remove_dup(a: list) -> None:
    """Remove duplicate strings from a given list.

    :param a: List to iterate through
    """
    i = 0
    while i < len(a):
        j = i + 1
        while j < len(a):
            if a[i] == a[j]:
                del a[j]
            else:
                j += 1
        i += 1


# Open needed files and bind to variables
folder = 'U:\\Joshua\\Work-Stuff\\Mimecast\\'

root = tk.Tk()
my_filetypes = [('csv files', '.csv')]


# Ask the user to select a single file name.
in_file = filedialog.askopenfilename(parent=root,
                                     initialdir=os.getcwd(),
                                     title="Please select a mimecast list file:",
                                     filetypes=my_filetypes)
# initialize input list
rows = []
timestr = time.strftime("%Y%m%d-%H%M%S")
headers = ['Email Address',
           'Policy (block, exempt, quarantine)',
           'Comment (optional)',
           ]


# populate input list from csv rows.
with open(in_file, encoding='utf-8', newline='') as csv_file:
    for i in range(3):
        next(csv_file)
    reader = csv.reader(csv_file)
    for row in reader:
        rows.append(row)

list_size = len(rows)
new_list = []
final_list = []
top_domains = ['google.com', 'gmail.com', 'yahoo.com',
            'outlook.com', 'aol.com', 'hotmail.com',
            'bellsouth.net', 'mail.com', 'microsoft.com',
            ]


# Only move unanimous block/allow to new list.
i = 0
while i < len(rows):
    j = i + 1
    conflict_detect = 0
    unanimous = 0
    while j < len(rows):
        if rows[i][0] == rows[j][0]:
            if rows[i][4] != rows[j][4]:
                conflict_detect += 1
                del rows[j]
                continue
            else:
                unanimous += 1
                del rows[j]
        else:
            j += 1

    if conflict_detect == 0 and unanimous > 0:
        if list_size <= 10000 and unanimous > 0:
            new_list.append(rows[i])
        elif 20000 >= list_size > 10000 and unanimous > 1:
            new_list.append(rows[i])
        elif 30000 >= list_size > 20000 and unanimous > 2:
            new_list.append(rows[i])
        elif 40000 >= list_size > 30000 and unanimous > 3:
            new_list.append(rows[i])
        elif 50000 >= list_size > 40000 and unanimous > 4:
            new_list.append(rows[i])
        elif list_size > 50000 and unanimous > 5:
            new_list.append(rows[i])
        i += 1
    else:
        i += 1

# Format unanimous list for cuda.
for i in new_list:
    if i[0] == '<>':
        continue
    elif i[0] in top_domains:
        continue
    else:
        if i[4] == 'Block':
            final_list.append([f"{i[0]}", "block", "from Mimecast"])
        elif i[4] == 'Permit':
            final_list.append([f"{i[0]}", "exempt", "from Mimecast"])
        elif i[4] == 'Quarantine':
            final_list.append([f"{i[0]}", "quarantine", "from Mimecast"])

# Remove any remaining true duplicates
remove_dup(final_list)

# Out file.
# TODO getcwd and output file there so anyone can run this easily.
out_file = f'{folder}processed-{timestr}.csv'
with open(out_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file)
    writer.writerow(headers)
    writer.writerows(final_list)

#TODO add simple tkinter window to maybe ask company name and edit out file etc.