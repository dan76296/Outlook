import argparse
import csv
import logging
import os
import sys
import win32com.client
import win32com

_logger = logging.getLogger(__name__)

from outlook import Outlook

Ol = Outlook()

Ol.connect_outlook()
Ol.get_accounts()
Ol.set_variables("Creepy Crawlies WiFi Login Event", "dantestscript12@ol.com")

for account in Ol.accounts:
    if account.DisplayName == Ol.email_address:
        Ol.get_inbox()
        Ol.set_archive_folder('Archived Wifi')
        Ol.parse_inbox()
        

# writer.writerow({
#             #         'login': results['Login'],
#             #         'hotspot_id': results['Hotspot ID'],
#             #         'name': results['Name'],
#             #         'email': results['Email'], 
#             #         'how': results['How did you find us'],
#             #         'browser': results['Browser'],
#             #         'mac': results["MAC Adress"] 
#             #         })