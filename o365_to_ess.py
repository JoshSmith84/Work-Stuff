f = open("C:\\temp\\ESSList.txt", "a")

AllowedSenders = ' '
AllowedSenderDomains = ' '
BlockedSenders = ' '


newAllowedList = ''
newBlockedList = ''

for i in AllowedSenders:
    if i == ' ':
        i = ',exempt,from O365\n'
    newAllowedList += i

for i in AllowedSenderDomains:
    if i == ' ':
        i = ',exempt,from O365\n'
    newAllowedList += i

for i in BlockedSenders:
    if i == ' ':
        i = ',block,from O365\n'
    newBlockedList += i



print(newAllowedList, file=f)
print(newBlockedList, file=f)
f.close()