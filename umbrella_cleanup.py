# umbrella_cleanup.py: script to drive browser to go through all clients
# and see if auto-delete after 30 days is enabled.
# If not enabled, enable. post status in csv to report on progress.
# Since it will be checking and only enabling those disabled,
# that will make this easier during debug process.

# Author: Josh Smith

import time
import sys
import shelve
import csv
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def run_auto_del(cust: list):
    global out_list
    global browser
    temp_list = []
    customer = cust[0]
    org_id = cust[1]
    temp_list.append(customer)
    temp_list.append(org_id)

    cust_url = f'https://dashboard.umbrella.com/o/' \
               f'{org_id}/#/deployments/core/roamingdevices/settings'

    browser.get(cust_url)

    body = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body')))
    webdriver.ActionChains(browser).move_to_element(body)
    time.sleep(3)
    old_auto_del = browser.find_element(By.XPATH, '//*[@id="dashx-shim-content"]/div/div/div/div/div[1]/div/div/div/div[2]/div/div/div[1]/div[3]/div[1]/div')
    old_ad = browser.find_element(By.XPATH, '//*[@id="dashx-shim-content"]/div/div/div/div/div[1]/div/div/div/div[2]/div/div/div[2]/div[3]/div[1]/div')
    old_vpn = browser.find_element(By.XPATH, '//*[@id="dashx-shim-content"]/div/div/div/div/div[1]/div/div/div/div[2]/div/div/div[3]/div[3]/div[1]/div')
    old_auto_del_status = old_auto_del.get_attribute("aria-checked")
    old_ad_status = old_ad.get_attribute("aria-checked")
    old_vpn_status = old_vpn.get_attribute("aria-checked")
    # auto_del_button = browser.find_element(By.XPATH,
    #                                        '//*[@id="dashx-shim-content"]/div/div/div/div/div[1]/div/div/div/div[2]/div/div/div[1]/div[3]/div[1]/div/div[2]')

    temp_list.append(old_auto_del_status)
    # print(f'old ad status: {old_ad_status}')
    # print(f'old vpnstatus: {old_vpn_status}')

    if old_auto_del_status == 'false':
        print(f'{company}: false')
        webdriver.ActionChains(browser).move_to_element(body)
        webdriver.ActionChains(browser).move_to_element(old_auto_del).click(
            old_auto_del).perform()
        webdriver.ActionChains(browser).move_to_element(body)
        time.sleep(1)
        new_auto_del = browser.find_element(By.XPATH, '//*[@id="dashx-shim-content"]/div/div/div/div/div[1]/div/div/div/div[2]/div/div/div[1]/div[3]/div[1]/div')
        # new_ad = wait.until(EC.visibility_of_element_located((By.XPATH,
        #                                                       '//*[@id="dashx-shim-content"]/div/div/div/div/div[1]/div/div/div/div[2]/div/div/div[2]/div[3]/div[1]/div')))
        # new_vpn = wait.until(EC.visibility_of_element_located((By.XPATH,
        #                                                        '//*[@id="dashx-shim-content"]/div/div/div/div/div[1]/div/div/div/div[2]/div/div/div[3]/div[3]/div[1]/div')))
        new_auto_del_status = new_auto_del.get_attribute("aria-checked")
        # new_ad_status = new_ad.get_attribute("aria-checked")
        # new_vpn_status = new_vpn.get_attribute("aria-checked")
        temp_list.append(new_auto_del_status)
        # print(f"new auto Del status: {new_auto_del_status}")
        # print(f'new ad status: {new_ad_status}')
        # print(f'new vpnstatus: {new_vpn_status}')
    else:
        print(f'{company}: true')
        temp_list.append(' ')

    temp_list.append(old_ad_status)
    temp_list.append(old_vpn_status)
    time.sleep(0.5)
    back = browser.find_element(By.XPATH,
                                '/html/body/div[1]/div[2]/div[2]/div/div/div[1]')
    back.click()
    cust_man = wait.until(EC.visibility_of_element_located((
        By.CLASS_NAME, "card-wrapper___3XgIv")))

    out_list.append(temp_list)


primary_url = \
    'https://dashboard.umbrella.com/msp/659637#customermanagement/customers'

folder = 'U:\\Joshua\\Dropbox\\Dropbox\\Python\\Work Stuff\\'
in_file = f'{folder}umbrella_companies.csv'
out_file = f'{folder}umbrella_auto_del_complete.csv'
in_list = []
out_list = []
headers = ['Company', 'Org ID', 'AutoDelete Before', 'AutoDelete After', 'AD Status', 'VPN Status']

with shelve.open('umbrella') as db:
    ur = db['umbrella_u']
    pw = db['umbrella_p']

with open(in_file, encoding='utf-8', newline='') as csv_file:
    reader = csv.reader(csv_file)
    for row in reader:
        in_list.append(row)

browser = webdriver.Chrome(ChromeDriverManager().install())
browser.get(primary_url)
try:
    user_elem = browser.find_element(By.ID, 'username')
except NoSuchElementException:
    sys.exit('Could not find Log In Button. Exiting Program')

user_elem.send_keys(ur, Keys.TAB, pw, Keys.ENTER)

wait = WebDriverWait(browser, 20)
cust_man = wait.until(EC.visibility_of_element_located((
    By.CLASS_NAME, "card-wrapper___3XgIv")))

for company in in_list:
    run_auto_del(company)

with open(out_file, 'w', encoding='utf-8', newline='') as output_file:
    writer = csv.writer(output_file, quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(headers)
    writer.writerows(out_list)
