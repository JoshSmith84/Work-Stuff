import csv


def remove_dup_1(a: list) -> None:
    """Remove duplicate strings from a list within a list.
    Only focusing on the first item though as it is the 'sender' email
    of each and the only duplicate I'm concerned with

    :param a:The list to iterate through
    """
    i = 0
    while i < len(a):
        j = i + 1
        while j < len(a):
            if a[i][0] == a[j][0]: # added the index of 0 for each
                del a[j]
            else:
                j += 1
        i += 1


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
f1 = open("C:\\temp\\USBoutlist.txt", "a")

f2 = open("U:\\Joshua\\Dropbox\\Dropbox\\Documents\\USBinlist.csv", "r")

# initialize input list
rows = []

# populate input list from csv rows. (makes list of lists)
with f2 as csvfile:
    csvreader = csv.reader(csvfile)
    for row in csvreader:
        rows.append(row)

for i in rows:
    if len(i) < 3:
        rows.remove(i)
    elif i == False:
        rows.remove(i)
    elif i[0] == 'Sender':
        rows.remove(i)
    elif i[0] == '<>':
        rows.remove(i)
    else:
        continue

# remove duplicates before processing
remove_dup_1(rows)

# initialize out list
newlist = []

# iterate though in-list and format out-list
for i in rows:
    if len(i) < 3:
        continue
    elif i[3] == 'Block':
        newlist.append(f"{i[0]},block,from Mimecast")
    elif i[3] == 'Permit':
        newlist.append(f"{i[0]},exempt,from Mimecast")

# Barracuda found duplicates in the first run, added this second check
# remove_dup(newlist)

# print each Barracuda-formatted item onto a new line and add to output file
for i in newlist:
    print(i, file=f1)
f1.close()
f2.close()