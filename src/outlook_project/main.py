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
Ol.set_variables("Creepy Crawlies WiFi Login Event", "dantestscript12@outlook.com")

for account in Ol.accounts:
    if account.DisplayName == Ol.email_address:
        Ol.get_inbox()
        Ol.set_archive_folder('Archived Wifi')
        Ol.parse_inbox()
        Ol.write_data_to_csv()
        