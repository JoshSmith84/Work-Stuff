import csv

# declare variables
csv_intune = 'C:\\temp\\Causeway-Intune.csv'
csv_ncentral = 'C:\\temp\\Causeway-ncentral.csv'
intune_dict = {}
ncentral_dict = {}
missing_dict = {}
causeway_missing = "C:\\temp\\causeway_missing.txt"

# open intune csv and pull what I needed
with open(csv_intune, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        intune_dict[row[0].upper()] = row[7]

# open ncentral csv and pull what I needed
with open(csv_ncentral, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        ncentral_dict[row[0].upper()] = row[2]

# populate new dictionary with differences
for key, value in intune_dict.items():
    if key not in ncentral_dict:
        missing_dict[key] = value
    else:
        continue

# write output file
with open(causeway_missing, 'w') as file:
    for key, value in missing_dict.items():
        file.write(f"{key} : {value}\n")