import argparse
import csv
import logging
import os
import sys
import win32com.client
import win32com

_logger = logging.getLogger(__name__)
_logger.setLevel(logging.DEBUG)
# create file handler which logs even debug messages
fh = logging.FileHandler('outlook.log')
fh.setLevel(logging.DEBUG)
# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.ERROR)
# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)
# add the handlers to the logger
_logger.addHandler(fh)
_logger.addHandler(ch)

class Outlook:

    def __init__(self):
        self._logger = _logger
        self.field_list = ['login', 'hotspot_id', 'name', 'email', 'how', 'browser', 'mac']
        pass
    
    def set_variables(self, wifi_subject, email_address):
        self.wifi_subject = wifi_subject
        self.email_address = email_address
        self._logger.info("Set subject and email address variables.")

    def connect_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
            self._logger.info("Connected to Outlook successfully")
        except:
            self._logger.error("An unexpected error occurred whilst connecting to Outlook")
        
    def get_accounts(self):
        try:
            self.accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
            self._logger.info("I have found %s accounts." % len(self.accounts))
        except:
            self._logger.error("There are no Outlook accounts on this machine.")

    def get_inbox(self):
        self.inbox = self.outlook.GetDefaultFolder(6)
        if self.inbox is not None:
            self._logger.info("Retrieved the inbox of current account.")
        else:
            self._logger.error("There was a problem retrieving the inbox of current account.")

    def set_archive_folder(self, folder):
        self.archive = self.inbox.Folders(folder)
        self._logger.info("Set the archive folder to %s" % folder)
    
    def archive_message(self, message, results):
        try:
            message.move(self.archive)
            self._logger.info("Successfully archived: %s" % results)
        except:
            self._logger.error("There was an error archiving the message")
            self._logger.info(results)

    def parse_inbox(self):
        messages = self.inbox.Items
        self.data_list = []
        for message in messages:
            results = {}
            if message.subject == self.wifi_subject:
                lines = (message.body).split('\r\n')
                lines_stripped = [line.strip() for line in lines if line != '']

                for item in lines_stripped:

                    if 'MAC Adress' in item:
                        values = item.split(" ")
                        values = " ".join(values[:2]), values[2]
                        
                    else:
                        values = item.split(':')

                    if len(values) > 1:
                        results[values[0]] = values[1].strip(' ')
                
                self.data_list.append(results)
                self.archive_message(message, results)

    def write_data_to_csv(self):
        with open('wifi_information.csv','a',newline='') as f:
            writer = csv.writer(f)
            rows = []
            for entry in self.data_list:

                login = entry['Login']
                hotspot_id = entry['Hotspot ID']
                name = entry['Name']
                email = entry['Email']
                how = entry['How did you find us']
                browser = entry['Browser']
                mac_address = entry['MAC Adress:']

                row =  [login, hotspot_id, name, email, how, browser, mac_address]
                rows.append(row)
                
            writer.writerows(rows)
