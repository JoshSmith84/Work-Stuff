f = open("C:\\temp\\RiseESSList.txt", "w")

AllowedSenders = '
AllowedSenderDomains = '
BlockedSenders = '


newAllowedList = ''
newBlockedList = ''

for i in AllowedSenders:
    if i == ' ':
        i = ',exempt,from Appriver\n'
    newAllowedList += i

for i in AllowedSenderDomains:
    if i == ' ':
        i = ',exempt,from Appriver\n'
    newAllowedList += i

for i in BlockedSenders:
    if i == ' ':
        i = ',block,from Appriver\n'
    newBlockedList += i



print(newAllowedList, file=f)
print(newBlockedList, file=f)
f.close()