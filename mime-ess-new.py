import csv
import tkinter as tk
from tkinter import filedialog
import os
import time
from remove_dup import remove_dup



# Open needed files and bind to variables
folder = 'U:\\Joshua\\Work-Stuff\\Mimecast\\'

root = tk.Tk()
my_filetypes = [('csv files', '.csv')]


# Ask the user to select a single file name.
in_file = filedialog.askopenfilename(parent=root,
                                     initialdir=os.getcwd(),
                                     title="Please select an input file:",
                                     filetypes=my_filetypes)
# initialize input list
rows = []
domain_rows = []
timestr = time.strftime("%Y%m%d-%H%M%S")
headers = ['Email Address',
           'Policy (block, exempt, quarantine)',
           'Comment (optional)',
           ]


# populate input list from csv rows. (makes list of lists)
with open(in_file, encoding='utf-8', newline='') as csv_file:
    for i in range(3):
        next(csv_file)
    reader = csv.reader(csv_file)
    for row in reader:
        rows.append(row)

# init new list and edit to cuda out
# (TODO need to re-order this for more efficient script)
new_list = []

for i in rows:
    if i[4] == 'Quarantine' or len(i[0]) > 40:
        continue
    elif i[0] == '<>':
        continue
    else:
        if i[4] == 'Block':
            new_list.append([f"{i[0]}", "block", "from Mimecast"])
        elif i[4] == 'Permit':
            new_list.append([f"{i[0]}", "exempt", "from Mimecast"])

final_list = []

# Only move unanimous block/allow to final list.
# TODO clean this up.
#  No need to delete from new_list since we're not outputting that
i = 0
while i < len(new_list):
    j = i + 1
    d_orig = 0
    while j < len(new_list):
        if new_list[i][0] == new_list[j][0]:
            if new_list[i][1] != new_list[j][1]:
                d_orig += 1
                del new_list[j]
                continue
            else:
                final_list.append(new_list[j])
                del new_list[j]
        else:
            j += 1
    if d_orig == 1:
        del new_list[i]
        continue
    else:
        i += 1

top_doms = ['google.com', 'gmail.com', 'yahoo.com',
            'outlook.com', 'aol.com', 'hotmail.com',
            'bellsouth.net', 'mail.com', 'microsoft.com',
            ]

# Remove any remaining true duplicates
remove_dup(final_list)

# Out file.
# TODO Could just overwrite in file and then could wrap in exe for others.
out_file = f'{folder}processed-{timestr}.csv'
with open(out_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file)
    writer.writerow(headers)
    writer.writerows(final_list)
