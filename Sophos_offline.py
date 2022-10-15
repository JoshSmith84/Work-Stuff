#! python3
# sophos_offline.py - script I made specifically
# to see who is missing sophos and who has been offline over 45 days.

import csv

old_stuff = "C:\\temp\\Wag\\offline45d_9-26-2022.csv"
sophos_stuff = "C:\\temp\\Wag\\noSophos_9-26-2022.csv"

old_list = []
sophos_list = []
both_list = []
fresh_nosophos = []
both_csv = "C:\\temp\\Wag\\offline45d+sophos-missing.csv"
remaining = "C:\\temp\\Wag\\no-sophos-remaining.csv"

# open ncentral csv1 and pull what I needed
with open(old_stuff, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        old_list.append(row)

# open ncentral csv2 and pull what I needed
with open(sophos_stuff, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        sophos_list.append(row)

for row in sophos_list:
    if row in old_list:
        both_list.append(row)
for row in sophos_list:
    if row not in old_list:
        fresh_nosophos.append(row)

with open(both_csv, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(headers)
    writer.writerows(both_list)

with open(remaining, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(headers)
    writer.writerows(fresh_nosophos)
