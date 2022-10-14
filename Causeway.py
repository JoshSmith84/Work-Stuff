import csv

# declare variables
csv_tune = 'C:\\temp\\Causeway-tune.csv'
csv_central = 'C:\\temp\\Causeway-central.csv'
tune_dict = {}
central_dict = {}
missing_dict = {}
causeway_missing = "C:\\temp\\causeway_missing.txt"

# open intune csv and pull what I needed
with open(csv_tune, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        tune_dict[row[0].upper()] = row[7]

# open ncentral csv and pull what I needed
with open(csv_central, encoding='utf-8', newline='') as csv_file:
    headers = csv_file.readline().strip('\n').split(',')
    reader = csv.reader(csv_file)
    for row in reader:
        central_dict[row[0].upper()] = row[2]

# populate new dictionary with differences
for key, value in tune_dict.items():
    if key not in central_dict:
        missing_dict[key] = value
    else:
        continue

# write output file
with open(causeway_missing, 'w') as file:
    for key, value in missing_dict.items():
        file.write(f"{key} : {value}\n")