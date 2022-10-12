#! python3
# Clean up Ncentral device export
# Author: Josh Smith

import csv
import tkinter as tk
from tkinter import filedialog
import os


application_window = tk.Tk()
new_list = []
my_filetypes = [('csv files', '.csv')]


# Ask the user to select a single file name.
in_file = filedialog.askopenfilename(parent=application_window,
                                    initialdir=os.getcwd(),
                                    title="Please select an input file:",
                                    filetypes=my_filetypes)


# open device report csv and pull only what is needed
with open(in_file, encoding='utf-8', newline='') as csv_file:
    line1 = csv_file.readline().strip('\n').split(',')
    line2 = csv_file.readline().strip('\n').split(',')
    line3 = csv_file.readline().strip('\n').split(',')
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    del headers[7:]
    del headers[5]
    del headers[2:4]
    for row in reader:
        del row[7:]
        del row[5]
        del row[2:4]
        new_list.append(row)

#open input file and write output (overwrite)
with open(in_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(headers)
    writer.writerows(new_list)

