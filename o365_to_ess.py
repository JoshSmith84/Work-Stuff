import csv


def re_organize(raw: str, li: list):
    sender = ''
    for char in raw:
        if char == ' ':
            li.append(sender)
            sender = ''
            continue
        else:
            sender += char


def set_emails(in_list: list, action: str):
    global email_out_list
    for i in in_list:
        temp_list = []
        temp_list.append(i)
        temp_list.append(action)
        temp_list.append('From O365')
        email_out_list.append(temp_list)


def set_ips(in_list: list, action: str):
    global ip_out_list
    for i in in_list:
        temp_list = []
        temp_list.append(i)
        temp_list.append('255.255.255.255')
        temp_list.append(action)
        temp_list.append('From O365')
        ip_out_list.append(temp_list)


folder = 'C:\\temp\\'
out_file = f'{folder}email_ess.csv'
out_file2 = f'{folder}ip_ess.csv'

allowed_senders_raw = ''
allowed_domains_raw = ''
blocked_senders_raw = ''
blocked_domains_raw = ''
blocked_ips_raw = ''

allowed_senders_list = []
allowed_domains_list = []
blocked_senders_list = []
blocked_domains_list = []
blocked_ips_list = []

email_headers = ['Email Address', 'Policy (block, exempt, quarantine)', 'Comment (optional)']
ip_headers = ['IP Address', 'Netmask', 'Policy (block, exempt, quarantine)', 'Comment (optional)']

re_organize(allowed_senders_raw, allowed_senders_list)
re_organize(allowed_domains_raw, allowed_domains_list)
re_organize(blocked_senders_raw, blocked_senders_list)
re_organize(blocked_domains_raw, blocked_domains_list)
re_organize(blocked_ips_raw, blocked_ips_list)

email_out_list = []
ip_out_list = []
set_emails(allowed_senders_list, 'exempt')
set_emails(allowed_domains_list, 'exempt')
set_emails(blocked_senders_list, 'block')
set_emails(blocked_domains_list, 'block')
set_ips(blocked_ips_list, 'block')

with open(out_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(email_headers)
    writer.writerows(email_out_list)

with open(out_file2, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(ip_headers)
    writer.writerows(ip_out_list)

