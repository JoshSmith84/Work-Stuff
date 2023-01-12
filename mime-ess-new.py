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

new_list = []
final_list = []
top_domains = ['google.com', 'gmail.com', 'yahoo.com',
            'outlook.com', 'aol.com', 'hotmail.com',
            'bellsouth.net', 'mail.com', 'microsoft.com',
            ]


# Only move unanimous block/allow to final list.
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
    if conflict_detect > 0:
        del rows[i]
    elif conflict_detect == 0 and unanimous > 0:
        new_list.append(rows[i])
        i += 1
    else:
        i += 1

# Format unanimous list for cuda.
for i in new_list:
    if i[4] == 'Quarantine' or len(i[0]) > 40:
        continue
    elif i[0] == '<>':
        continue
    elif i[0] in top_domains:
        continue
    else:
        if i[4] == 'Block':
            final_list.append([f"{i[0]}", "block", "from Mimecast"])
        elif i[4] == 'Permit':
            final_list.append([f"{i[0]}", "exempt", "from Mimecast"])

# Remove any remaining true duplicates
remove_dup(final_list)

# Out file.
# TODO Could just overwrite in file and then could wrap in exe for others.
out_file = f'{folder}processed-{timestr}.csv'
with open(out_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file)
    writer.writerow(headers)
    writer.writerows(final_list)
