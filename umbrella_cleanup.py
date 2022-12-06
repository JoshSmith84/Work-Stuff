# umbrella_cleanup.py: script to drive browser to go through all clients
# and see if auto-delete after 30 days is enabled.
# If not enabled, enable. post status in csv to report on progress.
# Since it will be checking and only enabling those disabled,
# that will make this easier during debug process.

# Author: Josh Smith


from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options

#TODO define setting function. This will be in specific client portal.
# Click roaming computers-> Settings ->
# check if auto-delete is enabled, append list for csv row wwhen done


#TODO Need to do this TODO first: load site, wait for user to input creds.
# Go through and record all org IDs.
# populate them to a dictionary {company name, org ID}

#TODO iterate through the dictionary using the org IDs to open each
# company spcific portal, then function to do the setting.
# Pass the key as company name for the appending of list data for csv report.
