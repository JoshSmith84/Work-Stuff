f = open("C:\\temp\\ESSList.txt", "a")

AllowedIPs = '54.174.209.165 '


BlockedIPs=


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