#-------------------------------------------------------------------------------
# Name:        ExtractAccountsAndContactsFromAccess
# Purpose:
#
# Author:      esteban
#
# Created:     30/04/2013
# Copyright:   (c) esteban 2013
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#'Title', 'FirstName', 'MiddleName', 'LastName', 'Suffix', 'Company', 'Department', 'JobTitle', 'BusinessStreet', 'BusinessStreet2', 'BusinessStreet3', 'BusinessCity', 'BusinessState', 'BusinessPostalCode', 'BusinessCountryRegion', 'HomeStreet', 'HomeStreet2', 'HomeStreet3', 'HomeCity', 'HomeState', 'HomePostalCode', 'HomeCountryRegion', 'OtherStreet', 'OtherStreet2', 'OtherStreet3', 'OtherCity', 'OtherState', 'OtherPostalCode', 'OtherCountryRegion', 'AssistantsPhone', 'BusinessFax', 'BusinessPhone', 'BusinessPhone2', 'Callback', 'CarPhone', 'CompanyMainPhone', 'HomeFax', 'HomePhone', 'HomePhone2', 'ISDN', 'MobilePhone', 'OtherFax', 'OtherPhone', 'Pager', 'PrimaryPhone', 'RadioPhone', 'TTYTDDPhone', 'Telex', 'Account', 'Anniversary', 'AssistantsName', 'BillingInformation', 'Birthday', 'BusinessAddressPOBox', 'Categories', 'Children', 'DirectoryServer', 'EmailAddress', 'EmailType', 'EmailDisplayName', 'Email2Address', 'Email2Type', 'Email2DisplayName', 'Email3Address', 'Email3Type', 'Email3DisplayName', 'Gender', 'GovernmentIDNumber', 'Hobby', 'HomeAddressPOBox', 'Initials', 'InternetFreeBusy', 'Keywords', 'Language1', 'Location', 'ManagersName', 'Mileage', 'Notes', 'OfficeLocation', 'OrganizationalIDNumber', 'OtherAddressPOBox', 'Priority', 'Private', 'Profession', 'ReferredBy', 'Sensitivity', 'Spouse', 'User1', 'User2', 'User3', 'User4', 'WebPage']'Title', 'FirstName', 'MiddleName', 'LastName', 'Suffix', 'Company', 'Department', 'JobTitle', 'BusinessStreet', 'BusinessStreet2', 'BusinessStreet3', 'BusinessCity', 'BusinessState', 'BusinessPostalCode', 'BusinessCountryRegion', 'HomeStreet', 'HomeStreet2', 'HomeStreet3', 'HomeCity', 'HomeState', 'HomePostalCode', 'HomeCountryRegion', 'OtherStreet', 'OtherStreet2', 'OtherStreet3', 'OtherCity', 'OtherState', 'OtherPostalCode', 'OtherCountryRegion', 'AssistantsPhone', 'BusinessFax', 'BusinessPhone', 'BusinessPhone2', 'Callback', 'CarPhone', 'CompanyMainPhone', 'HomeFax', 'HomePhone', 'HomePhone2', 'ISDN', 'MobilePhone', 'OtherFax', 'OtherPhone', 'Pager', 'PrimaryPhone', 'RadioPhone', 'TTYTDDPhone', 'Telex', 'Account', 'Anniversary', 'AssistantsName', 'BillingInformation', 'Birthday', 'BusinessAddressPOBox', 'Categories', 'Children', 'DirectoryServer', 'EmailAddress', 'EmailType', 'EmailDisplayName', 'Email2Address', 'Email2Type', 'Email2DisplayName', 'Email3Address', 'Email3Type', 'Email3DisplayName', 'Gender', 'GovernmentIDNumber', 'Hobby', 'HomeAddressPOBox', 'Initials', 'InternetFreeBusy', 'Keywords', 'Language1', 'Location', 'ManagersName', 'Mileage', 'Notes', 'OfficeLocation', 'OrganizationalIDNumber', 'OtherAddressPOBox', 'Priority', 'Private', 'Profession', 'ReferredBy', 'Sensitivity', 'Spouse', 'User1', 'User2', 'User3', 'User4', 'WebPage']


#Account Category	Account Type	Account Source	Annual Revenue	Billing City	Billing Country	Billing Zip/Postal Code	Billing State/Province	Billing Street	Created By ID	Created Date	Account Description	Account Fax	Account ID	Industry	Deleted	Data.com Key	Jigsaw Company ID	Last Activity	Last Modified By ID	Last Modified Date	Master Record ID	Account Name	Employees	Owner ID	Parent Account ID	Account Phone	Shipping City	Shipping Country	Shipping Zip/Postal Code	Shipping State/Province	Shipping Street	SIC Description	System Modstamp	Account Type	Website

import csv
import sys
import datetime
import re
import argparse
import os

from collections import defaultdict
from datetime import datetime

sf_account_fields = ["Account Category", "Account Type", "Account Source", "Annual Revenue", "Billing City", "Billing Country", "Billing Zip/Postal Code", "Billing State/Province", "Billing Street", "Created By ID", "Created Date", "Account Description", "Account Fax", "Account ID", "Industry", "Deleted", "Data.com Key", "Jigsaw Company ID", "Last Activity", "Last Modified By ID", "Last Modified Date", "Master Record ID", "Account Name", "Employees", "Owner ID", "Parent Account ID", "Account Phone", "Shipping City", "Shipping Country", "Shipping Zip/Postal Code", "Shipping State/Province", "Shipping Street", "SIC Description", "System Modstamp", "Website"]
#sf_contact_fields = ["Account ID","Account Name","Assistant's Name","Asst. Phone","Birthdate","Categories","Community","Contact Type","Created By ID","Created Date","Department","Contact Description","Email","Email Bounced Date","Email Bounced Reason","Business Fax","First Name","Home Phone","Contact ID","Deleted","Data.com Key","Jigsaw Contact ID","Last Activity","Last Stay-in-Touch Request Date","Last Stay-in-Touch Save Date","Last Modified By ID","Last Modified Date","Last Name","Lead Source","Mailing City","Mailing Country","Mailing Zip/Postal Code","Mailing State/Province","Mailing Street","Master Record ID","Mobile Phone","Full Name","Other City","Other Country","Other Phone","Other Zip/Postal Code","Other State/Province","Other Street","Owner ID","Business Phone","Political Party","Project Association","Reports To ID","Salutation","System Modstamp","Title"]
sf_contact_fields = ["Account ID","Account Name","Assistant's Name","Asst. Phone","Birthdate","Categories","Community","Contact Type","Created By ID","Created Date","Department","Contact Description","Email","Email Bounced Date","Email Bounced Reason","Business Fax","First Name","Home Phone","Contact ID","Deleted","Data.com Key","Jigsaw Contact ID","Last Activity","Last Stay-in-Touch Request Date","Last Stay-in-Touch Save Date","Last Modified By ID","Last Modified Date","Last Name","Lead Source","Mailing City","Mailing Country","Mailing Zip/Postal Code","Mailing State/Province","Mailing Street","Master Record ID","Mobile Phone","Full Name","Other City","Other Country","Other Phone","Other Zip/Postal Code","Other State/Province","Other Street","Owner ID","Business Phone","Political Party","Project Association","Reports To ID","Salutation","System Modstamp","Title","Home City","Home Country","Home Zip/Postal Code","Home State/Province","Home Street"]
sf_contact_fields_count = len(sf_contact_fields)

access_fields = ["MainID","meTemp","meLName","meFName","meMInit","meTitle","meAff","lkSal1","lkSal3","lkSal2","lkSal4","meOAdd","lkOCity","lkOState","meOZip","meODPhone","meOPhone","meOFax","meCPhone","meCarPhone","mePager","meXPhone","meIntPhone","meIntFax","meEMail","meWeb","meHSpouse","meHAdd","lkHCity","lkHState","meHZip","meHPhone","meHFax","meHXPhone","meHEMail","meEOAdd","lkEOCity","lkEOState","meEOZip","meEOPhone","meEOFax","meEOEMail","meEOParty","meEODNumber","meEOChief","meEOSch","Categories"]

sf_for_access_field ={"Account ID":"","Account Name":"meAff","Assistant's Name":"meEOSch","Asst. Phone":"","Birthdate":"","Categories":"Categories","Community":"","Contact Type":"","Created By ID":"","Created Date":"","Department":"","Contact Description":"","Email":"meEMail","Email Bounced Date":"","Email Bounced Reason":"","Business Fax":"","First Name":"meFName","Direct Phone":"meODPhone","Home Phone":"meHPhone","Contact ID":"","Deleted":"","Data.com Key":"","Jigsaw Contact ID":"","International Phone":"meIntPhone","International Fax":"meIntFax","Last Activity":"","Last Stay-in-Touch Request Date":"","Last Stay-in-Touch Save Date":"","Last Modified By ID":"","Last Modified Date":"","Last Name":"meLName","Lead Source":"","Middle Name":"meMInit","Mailing City":"lkOCity","Mailing Country":"","Mailing Zip/Postal Code":"meOZip","Mailing State/Province":"lkOState","Mailing Street":"meOAdd","Master Record ID":"","Mobile Phone":"meCPhone","Full Name":"","Other City":"lkEOCity","Other Country":"","Other Phone":"meEOPhone","Other Fax":"meEOFax","Other Zip/Postal Code":"meEOZip","Other State/Province":"lkEOState","Other Street":"meEOAdd","Other Address Type":"","Other Email":"meEOEMail","Owner ID":"","Business Phone":"meOPhone","Business Fax":"meOFax","Political Party":"meEOParty","Project Association":"","Reports To ID":"","Salutation":"lkSal1","System Modstamp":"","Title":"meTitle","Website":"meWeb","Home City":"lkHCity","Home Country":"","Home Zip/Postal Code":"meHZip","Home State/Province":"lkHState","Home Street":"meHAdd","Home Fax":"meHFax","Personal Email":"meHEMail","Chief of Staff":"meEOChief","Scheduler":"meEOSch","District Number":"meEODNumber","Home City":"lkHCity","Home Country":"","Home Zip/Postal Code":"meHZip","Home State/Province":"lkHState","Home Street":"meHAdd"}

#def processAccessRowIntoSalesforceContact(data_row):
#	if row[LASTNAME] and row[FIRSTNAME]:
		
#	return


parser = argparse.ArgumentParser(description='Convert an Outlook CVS exported contacts list into separate Accounts and Contacts cvs file for import into Salesforce.')
parser.add_argument('-f', help='CSV file to process', dest='file')
parser.add_argument('-d', help='Where to save the output CSV files.', dest='output_dir')
parser.add_argument('-p', help='Prefix for output CSV files.', dest='prefix')
parser.add_argument('-c', help="Categories (comma separated) to add all records to", dest="categories")
parser.print_help()
args = parser.parse_args()

cats = args.categories
print(cats)
file = args.file
saveDirectory = args.output_dir

if not os.path.exists(file):
	print("ERROR: no file exists at '" + file + "'")
	exit()

if saveDirectory:
	if not os.path.exists(saveDirectory):
		print("ERROR: no such directory exists at '" + saveDirectory + "'")
		exit()

file = os.path.abspath(file)  #find absolute path to file
filename = os.path.basename(file) #find just the file name

prefix = args.prefix 
altprefix = prefix
now = datetime.now()
if not prefix:
	prefix = "SF" + now.strftime("%M%S")

altprefix = prefix + now.strftime("%M%S") + "_"
prefix = prefix + "_"

if not saveDirectory:
	saveDirectory = os.path.dirname(file)
	if not saveDirectory:
#		print("found save directory from file: " + saveDirectory)
#	else:
		saveDirectory = os.getcwd()
#		print("found save directory via cwd: " + saveDirectory)
else:
	saveDirectory = os.path.abspath(saveDirectory)
#	print("Made sure we had the right save directory: " + saveDirectory)

(root, ext) = os.path.splitext(filename)
if not '.csv' in ext:
	print("***OMG*** '" + filename + "' doesn't look like a CSV file, but I am going to try anyway.  Because I'm a trooper!  Don't steer me wrong!")

readCSVfile = file
now = datetime.now()
timestamp = now.strftime("%Y%m%d") 
accountfile = os.path.join(saveDirectory, prefix + 'Accounts-' + timestamp + ".csv")
if os.path.exists(accountfile):
	accountfile = os.path.join(saveDirectory, altprefix + 'Accounts-' + timestamp + ".csv")

contactfile = os.path.join(saveDirectory, prefix + 'Contacts-' + timestamp + ".csv")
if os.path.exists(contactfile):
	contactfile = os.path.join(saveDirectory, altprefix + 'Contacts-' + timestamp + ".csv")
#contactfile = saveDirectory + prefix + 'Contacts-' + timestamp  + ".csv"
badEmailContactFile = saveDirectory + prefix + 'BadEmailContacts-' + timestamp + ".csv"

print("accountfile: " + accountfile + "; contactfile: " + contactfile + " bademailcontactfile: " + badEmailContactFile)

FIRSTNAME = 'meFName'
LASTNAME = 'meLName'
COMPANY_FIELD = 'meAff'
EMAIL_FIELD = 'meEMail'

with open(readCSVfile, 'r') as csvfile:
	reader = csv.DictReader(csvfile, delimiter=',', quotechar='\"')
	accounts = defaultdict(str)
	categoriesList = []
	totalAccounts = 0
	crmAccounts = 0
	clientAccounts = 0
	vendorAccounts = 0
	
	csvcontact = open(contactfile, "w+")
	csvcontactwriter = csv.writer(csvcontact, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
	csvcontactwriter.writerow(sf_contact_fields)
	
	csvbademailfile = open(badEmailContactFile, "w+")
	csvbademailwriter = csv.writer(csvbademailfile, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
	csvbademailwriter.writerow(['FirstName', 'LastName', 'Company', 'Email'])
	
	for row in reader:
		totalAccounts+=1
		
		categories = ""
#		if 'Categories' in row:
#			categories = row['Categories']
#			categoriesList.append(categories)
						
		contactRowValuesDict = defaultdict(str)
		contactRowValues = []
		
		if row[LASTNAME] and row[FIRSTNAME]:
			for field in sf_contact_fields:
				if field in sf_for_access_field.keys():
					data = ""
					access_key = sf_for_access_field[field]
					if access_key in row.keys():
						data = row[access_key]
					contactRowValuesDict[field] = data

#			for key in sf_contact_fields: 
#				if key in contactRowValuesDict:  
#				else:
#					contactRowValues.append("")
			
			for field in sf_contact_fields:
				if field in contactRowValuesDict:
					if 'Email' in field and len('Email') == len(field):
						email = row[EMAIL_FIELD]
						if not re.match(r"[^@]+@[^@]+\.[^@^`]+", email) and email:
							contactRowValues.append("")
							print(row[LASTNAME] +" " +row[FIRSTNAME] + " - Bad E-mail! ('" + row[EMAIL_FIELD] + "'). Recovering by throwing away junk e-mail address!")
							csvbademailwriter.writerow([row[FIRSTNAME],row[LASTNAME],row[COMPANY_FIELD],row[EMAIL_FIELD]])
						else:
							contactRowValues.append(email)
					else:
						contactRowValues.append(contactRowValuesDict[field].strip())
				else:
					contactRowValues.append("")

			#sanity check
			row_fields_count = len(contactRowValues)
			if not (row_fields_count == sf_contact_fields_count):
				if row_fields_count > sf_contact_fields_count:
					print("row fields " + str(row_fields_count) + " > [" + str(sf_contact_fields_count) + "]")
				elif row_fields_count < sf_contact_fields_count:
					print("row fields " + str(row_fields_count) + " < [" + str(sf_contact_fields_count) + "]")
				print("******************")
				print(sf_contact_fields)
				print(contactRowValues)
					
			csvcontactwriter.writerow(contactRowValues)

			accountNameKey = row[COMPANY_FIELD]
			if accountNameKey:
				if accountNameKey in accounts:
					if not categories in accounts[accountNameKey]:
						accounts[accountNameKey] = ",".join([accounts[accountNameKey],categories])
				else:
					accounts[accountNameKey] = categories
		else:
			if not row[LASTNAME] and not row[FIRSTNAME]:
				print("Record has no First and Last Name. Can't add to database");
			else:
				print("Record with incomplete last name: '" + row[LASTNAME] + "', first name: '" + row[FIRSTNAME] + "'. Can't add to database.")
			
	
	print("Total Accounts: " + str(totalAccounts))

	csvaccount = open(accountfile, "w+")
	csvaccountwriter = csv.writer(csvaccount, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
	
	csvaccountwriter.writerow(sf_account_fields)
	
	for accountName in sorted(accounts.keys()):
#		print(accountName)
		accountRowValues = []
		for sf_account_field_name in sf_account_fields:
			if "Account Category" in sf_account_field_name:
				accountRowValues.append(accounts[accountName])
			elif "Account Name" in sf_account_field_name:
				accountRowValues.append(accountName)
			else:
				accountRowValues.append("")
		csvaccountwriter.writerow(accountRowValues)
	
#	categoriesList = sorted(set(categoriesList))
	
#	close(accountfile)
#	close(contactfile)
	
#Print the categories in our CSV file from Outlook
#	print("Categories: " + ', '.join(categoriesList))

	
#with open(accountfile, 'w+') as csvaccount:
#	csvaccountwriter = csv.write(csvaccount, delimiter=', ', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
#	header = ("\t".join(sf_account_fields)).strip()
#	header.replace("\t", ",")
#	csvaccountwriter = csv.writerow(header)

