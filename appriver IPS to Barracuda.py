f = open("C:\\temp\\RiseESSList.txt", "a")

AllowedIPs = '54.174.209.165 '


BlockedIPs= '167.89.13.221 167.89.67.124 167.89.83.148 167.89.95.233 209.85.222.66 '


newAllowedList = ''
newBlockedList = ''

for i in AllowedIPs:
    if i == ' ':
        i = ',255.255.255.255,exempt,from Appriver\n'
    newAllowedList += i

for i in BlockedIPs:
    if i == ' ':
        i = ',255.255.255.255,block,from Appriver\n'
    newBlockedList += i



print(newAllowedList, file=f)
print(newBlockedList, file=f)
f.close()