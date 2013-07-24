import csv
import sys
import datetime
import re
import argparse
import os
import glob

from collections import defaultdict
from datetime import datetime


account_fields = ["Account Category", "Account Type", "Account Source", "Annual Revenue", "Billing City", "Billing Country", "Billing Zip/Postal Code", "Billing State/Province", "Billing Street", "Created By ID", "Created Date", "Account Description", "Account Fax", "Account ID", "Industry", "Deleted", "Data.com Key", "Jigsaw Company ID", "Last Activity", "Last Modified By ID", "Last Modified Date", "Master Record ID", "Account Name", "Employees", "Owner ID", "Parent Account ID", "Account Phone", "Shipping City", "Shipping Country", "Shipping Zip/Postal Code", "Shipping State/Province", "Shipping Street", "SIC Description", "System Modstamp", "Website"]
contact_fields = ["Account ID","Account Name","Assistant's Name","Asst. Phone","Birthdate","Categories","Community","Contact Type","Created By ID","Created Date","Department","Contact Description","Email","Email Bounced Date","Email Bounced Reason","Business Fax","First Name","Home Phone","Contact ID","Deleted","Data.com Key","Jigsaw Contact ID","Last Activity","Last Stay-in-Touch Request Date","Last Stay-in-Touch Save Date","Last Modified By ID","Last Modified Date","Last Name","Lead Source","Mailing City","Mailing Country","Mailing Zip/Postal Code","Mailing State/Province","Mailing Street","Master Record ID","Mobile Phone","Full Name","Other City","Other Country","Other Phone","Other Zip/Postal Code","Other State/Province","Other Street","Owner ID","Business Phone","Political Party","Project Association","Reports To ID","Salutation","System Modstamp","Title","Home City","Home Country","Home Zip/Postal Code","Home State/Province","Home Street"]

contact_file_sig = "*_Contact*.csv"
account_file_sig = "*_Account*.csv"

EMAIL = "Email"
FIRSTNAME = "First Name"
LASTNAME = "Last Name"
ACCOUNTNAME = "Account Name"

parser = argparse.ArgumentParser(description='Combine Saleforce-formatted Account and Contact files.')
parser.add_argument('-d', required=True, help='Directory with Account and Contact files.  This is also the destination of the resulting master Account and master Contact files.', dest='working_dir')
parser.add_argument('-e', help="Destination Directory for output", dest='dest_dir')
parser.add_argument('-p', help='Prefix for master Account and Contact files generated.', dest='prefix')
parser.print_help()
args = parser.parse_args()

working_dir = args.working_dir
prefix = args.prefix
dest_dir = args.dest_dir

if working_dir:
	if not os.path.exists(working_dir):
		print("ERROR: no such directory exists at '" + working_dir + "'")
	working_dir = os.path.normpath(working_dir)
	
if dest_dir:
	if not os.path.exists(dest_dir):
		print("ERROR: no such (destination) directoy exists at '" + dest_dir + "'")
	dest_dir = os.path.normpath(dest_dir)
else:
	dest_dir = working_dir
	
now = datetime.now()
timestamp = now.strftime("%M%S")

master_acct_file_name = timestamp + "Master Account.csv"
master_acct_file = os.path.join(dest_dir, master_acct_file_name)

if not os.path.exists(master_acct_file):
	master_file = open(master_acct_file, "wb+")
	master_csv = csv.writer(master_file, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
	master_csv.writerow(account_fields)
	master_file.close()

def createNewAccountFileAtDestDir():
	file_name = timestamp + "TEMP Account.csv"
	file_name = os.path.join(dest_dir, file_name)
	new_account_file = open(file_name, "wb+")
	new_csv = csv.writer(new_account_file, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
	new_csv.writerow(account_fields)
	new_account_file.close()
	return file_name
	
def contactRowsMatch(mrow, frow):
	if mrow[EMAIL] == frow[EMAIL]:
		if mrow[LASTNAME] == frow[LASTNAME]:
			if mrow[FIRSTNAME] == frow[FIRSTNAME]:
				return True
	return False

def accountRowsMatch(mrow, frow):
	if not ACCOUNTNAME in mrow:
		return False
		
	if not ACCOUNTNAME in frow:
		return False
		
	if mrow[ACCOUNTNAME] == frow[ACCOUNTNAME]:
		return True
	return False
	
def rowWithAccountNameValue(row):
	accountRowValues = []
	if ACCOUNTNAME in row:
		for header in account_fields:
			if header == ACCOUNTNAME:
				accountRowValues.append(row[ACCOUNTNAME])
			else:
				accountRowValues.append("")
	return accountRowValues
	
def writeMultipleRowValuesToFile(file, rows):
	filecsv = open(file, "r+b")
	fwriter = csv.writer(filecsv, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
	print "Rows we are writing to file '" + file + "': " + str(len(rows))
	fwriter.writerows(rows)
	filecsv.close()

def writeMultipleRowValuesToMasterFile(rows):
	mastercsv = open(master_acct_file, "r+b")
	mastercsv.seek(0,2)
	mwriter = csv.writer(mastercsv, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
	print "rows we think we have: " + str(len(rows))
	mwriter.writerows(rows)
	mastercsv.close()

			
def processFileForMasterAccountFile(file):
	if not file:
		return
	if not os.path.exists(file):
		print "ERROR: file '" + file + "' does not actually exist. It was just a dream."
		return
	print "Reading from: " + file
	
	filecsv = open(file, "r")
	freader = csv.DictReader(filecsv, delimiter=',', quotechar='\"')
	
	with open(master_acct_file, "r+b") as mastercsv:
		mreader = csv.DictReader(mastercsv, delimiter=',', quotechar='\"') 
		
		accountRowValuesToInsert = []

		row_count = 0
		for frow in freader:
			for mrow in mreader:
				if not accountRowsMatch(mrow,frow):
					accountRowValues = rowWithAccountNameValue(frow)
					accountRowValuesToInsert.append(accountRowValues)
					row_count = row_count + 1
			mastercsv.seek(0)
		
		print "rows we inserted: " + str(row_count)
		 #go to end of file
		mastercsv.close()
		
#	writeMultipleRowValuesToMasterFile(accountRowValuesToInsert)
	writeMultipleRowValuesToFile(createNewAccountFileAtDestDir(), accountRowValuesToInsert)
#		mastercsv = open(master_acct_file, "r+b")#
#		mastercsv.seek(0,2)
#		mwriter = csv.writer(mastercsv, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
#
#		print "rows we think we have: " + str(len(accountRowValuesToInsert))
#		holder = accountRowValuesToInsert
#		accountRowValuesToInsert = []
#		mwriter.writerows(holder)
#		mastercsv.close()
	filecsv.close()
	return	
		
acct_dir_match = os.path.join(working_dir, contact_file_sig)
cont_dir_match = os.path.join(working_dir, account_file_sig)

account_files = glob.glob(acct_dir_match) 
print "Account Directory Matching: " + acct_dir_match 
print "Account Files: " + "; ".join(account_files)
print "----------"
contact_files = glob.glob(cont_dir_match)
print "Contact Directory Matching: " + cont_dir_match + "\n"
print "Contact Files: " + "; ".join(contact_files)
print "----------"

for file in account_files:
	processFileForMasterAccountFile(file)