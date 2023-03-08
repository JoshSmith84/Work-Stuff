import csv




# Open needed files and bind to variables
folder = 'C:\\temp\\evaptech\\'
in_file = f'{folder}evap_in_list.csv'
out_file = f'{folder}evap_int_list.csv'

# initialize input list
rows = []
domain_rows = []
headers = ['Email Address',
           'Policy (block, exempt, quarantine)',
           'Comment (optional)',
           ]

# populate input list from csv rows. (makes list of lists)
with open(in_file, encoding='utf-8', newline='') as csv_file:
    reader = csv.reader(csv_file)
    for row in reader:
        rows.append(row)
print('done reading rows')
i = 0
while i < len(rows):
    j = i + 1
    while j < len(rows):
        if rows[i][0] == rows[j][0]:  # added the index of 0 for each
            if rows[i][2] != rows[j][2]:
                rows[i][2] = 'Quarantine'
            del rows[j]
        else:
            j += 1
    i += 1

new_list = []
for i in rows:
    if i[2] == 'Quarantine' or len(i[0]) > 40:
        continue
    else:
        if i[2] == 'Block':
            new_list.append([f"{i[0]}", "block", "from Mimecast"])
        elif i[2] == 'Permit':
            new_list.append([f"{i[0]}", "exempt", "from Mimecast"])

with open(out_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file)
    writer.writerow(headers)
    writer.writerows(new_list)