import csv

# declare variables
csv_filename1 = 'C:\\temp\\itarian.solalt.csv'
csv_filename2 = 'C:\\temp\\ncentral-solalt.csv'
itarian_dict = {}
ncentral_dict = {}
missing_dict = {}
solalt_missing = "C:\\temp\\causeway_missing.txt"

# open itarian csv and pull what I needed
with open(csv_filename1, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        itarian_dict[row[2]] = row[8]

# open ncentral csv and pull what I needed
with open(csv_filename2, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        ncentral_dict[row[0]] = row[2]

# populate new dictionary with differences
for key, value in itarian_dict.items():
    if key not in ncentral_dict:
        missing_dict[key] = value
    else:
        continue

# write output file
with open(solalt_missing, 'w') as file:
    for key, value in missing_dict.items():
        file.write(f"{key} : {value}\n")