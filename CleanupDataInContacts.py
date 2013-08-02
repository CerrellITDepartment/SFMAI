#!/bin/python

import csv
import sys
import datetime
import re
import argparse
import os
import glob

from collections import defaultdict
from datetime import datetime

contact_fields = ["Account ID","Account Name","Assistant's Name","Asst. Phone","Birthdate","Categories","Community","Contact Type","Created By ID","Created Date","Department","Contact Description","Email","Email Bounced Date","Email Bounced Reason","Business Fax","First Name","Home Phone","Contact ID","Deleted","Data.com Key","Jigsaw Contact ID","Last Activity","Last Stay-in-Touch Request Date","Last Stay-in-Touch Save Date","Last Modified By ID","Last Modified Date","Last Name","Lead Source","Mailing City","Mailing Country","Mailing Zip/Postal Code","Mailing State/Province","Mailing Street","Master Record ID","Mobile Phone","Full Name","Other City","Other Country","Other Phone","Other Zip/Postal Code","Other State/Province","Other Street","Owner ID","Business Phone","Political Party","Project Association","Reports To ID","Salutation","System Modstamp","Title","Home City","Home Country","Home Zip/Postal Code","Home State/Province","Home Street"]

HOME = "Home Phone"
MOBILE = "Mobile Phone"
OTHER = "Other Phone"
BUSINESS = "Business Phone"
FAX = "Business Fax"

HOMEZIP = "Home Zip/Postal Code"
MAILINGZIP = "Mailing Zip/Postal Code"
OTHERZIP = "Other Zip/Postal Code"

parser = argparse.ArgumentParser(description='Combine Saleforce-formatted Account and Contact files.')
parser.add_argument('-f', required=True, help='Contact file we are working with', dest='working_file')
parser.add_argument('-o', help="Outfile file to generate", dest='output_file')
parser.print_help()
args = parser.parse_args()

phone_re = re.compile(r'(\d{3})\D*(\d{3})\D*(\d{4})\D*(\d*)$', re.VERBOSE)
zip_re = re.compile(r'(\d{5})\D*(\d{4})*', re.VERBOSE)

def zipCleanup(zipcode):
	new_zip = zipcode
	if zipcode:
		search_results = zip_re.search(zipcode)
		if not search_results:
			print "ERROR (ZIP Cleanup): parsing '" + zipcode + "' failed - IGNORE and CONTINUE"
			return zipcode
		
		groups = search_results.groups()
		if len(groups) >= 1:
			new_zip = groups[0]
			if len(groups) == 2:
				if groups[1]:
					new_zip = new_zip + "-" + groups[1]
			else:
				new_zip = zipcode
	return new_zip
		
def phoneCleanup(phone_num):
	new_phone = phone_num
	if phone_num:
		search_results = phone_re.search(phone_num)
		if not search_results:
			print "ERROR (Phone # Cleanup): parsing '" + phone_num + "' failed - IGNORE and CONTINUE"
			return phone_num
	
		groups = search_results.groups()
		if len(groups) >= 3 and len(groups) <= 4:
			new_phone = "(" + groups[0] + ") " + groups[1] + "-" + groups[2]
			if len(groups) >= 4:
				if groups[3]:
					new_phone = new_phone + "-" + groups[3]
		else:
			new_phone = "-".join(phone_re.search(phone_num).groups())
			new_phone = new_phone[:-1]
		if not new_phone:
			print "ERROR: Clean up of '" + phone_num + "' => " + new_phone
	return new_phone

working_file = args.working_file
if working_file:
	if not os.path.exists(working_file):
		print("ERROR: no such file exists at '" + working_file + "'")
		exit()

output_file = args.output_file
if output_file:
	if os.path.exists(output_file):
		print("ERROR: no output file already exists at '" + output_file + "'")
		exit()
if not output_file:
	(path, tail) = os.path.split(working_file)
	if tail:
		(root, ext) = os.path.splitext(tail)
		if root:
			index = 0
			file_path = os.path.join(path, root + "_" + str(index) + ext)
			while os.path.exists(file_path):
				index = index + 1
				file_path = os.path.join(path, root + "_" + str(index) + ext)
				print file_path
			output_file = file_path

file = open(output_file, "wb")
writer = csv.DictWriter(file, contact_fields, delimiter=',', quotechar='\"', quoting=csv.QUOTE_MINIMAL)
writer.writeheader()

with open(working_file, "rb") as wfile:
	wf_reader = csv.DictReader(wfile, delimiter=',', quotechar='\"')	
	for row in wf_reader:
		row[MAILINGZIP] = zipCleanup(row[MAILINGZIP])
		row[HOMEZIP] = zipCleanup(row[HOMEZIP])
		row[OTHERZIP] = zipCleanup(row[OTHERZIP])
		
		row[BUSINESS] = phoneCleanup(row[BUSINESS])
		row[MOBILE] = phoneCleanup(row[MOBILE])		
		row[HOME] = phoneCleanup(row[HOME])		
		row[OTHER] = phoneCleanup(row[OTHER])		
		row[FAX] = phoneCleanup(row[FAX])
		writer.writerow(row)
	
	print "GREAT SUCCESS: Clean up of data for '" + working_file + "' complete and saved to '" + output_file + "'."