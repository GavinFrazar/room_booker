#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  room_booker.py
#  
#  rodericd5
#  
#  created 12/24/2017 

import time
import os.path
from datetime import datetime
from datetime import timedelta
from selenium import webdriver
import smtplib
import re
import urllib2
import imaplib
import email

UCSB_ADD = "@umail.ucsb.edu"
FORTNIGHT = 14
IMTP_ADD = "outlook.office365.com"
IMTP_PORT = 993
HARDCODED_DRIVER_LOCATION = '****************'

#enters if file does not already exist and prompts user to enter necessary info to run program
if (os.path.isfile('library_room.txt') == False):
	
	timeframe = raw_input("Please enter\n10 to book 10-12\n12 to book 12-2\n 2 to book 2-4\n 4 to book 4-6\n or 6 to book 6-8\n")
	timeframe_check = raw_input("Please enter the number again to confirm: ")
	while (timeframe != timeframe_check):
		print "Error the two numbers do not match. Please try again\n"
		timeframe = raw_input("Please enter\n10 to book 10-12\n12 to book 12-2\n 2 to book 2-4\n 4 to book 4-6\n or 6 to book 6-8\n")
		timeframe_check = raw_input("Please enter the number again to confirm: ")
		
	login_id_noadd = raw_input("Please enter your Net ID (without the .umail extension): ")
	login_id_noadd_check = raw_input("Please enter your Net ID again to confirm: ")
	while (login_id_noadd != login_id_noadd_check):
		print "Error the two Net ID's do not match. Please try again\n"
		login_id_noadd = raw_input("Please enter your Net ID (without the umail extension): ")
		login_id_noadd_check = raw_input("Please enter your Net ID again to confirm: ")
		
	login_pwd = raw_input("Please enter your Net ID password: ")
	login_pwd_check = raw_input("Please enter your password again to confirm: ")
	while (login_pwd != login_pwd_check):
		print "Error the two passwords do not match. Please try again\n"
		login_pwd = raw_input("Please enter your Net ID password: ")
		login_pwd_check = raw_input("Please enter your password again to confirm: ")
	
	#write the answers to the prompts line by line 
	f = open('library_room.txt', 'w')
	f.write(timeframe + '\n')
	f.write(login_id_noadd + '\n')
	f.write(login_pwd)
	f.close()

#opens file with necessary information	
f = open('library_room.txt', 'r')

#reads first line which determines times to book (need to get rid of '\n' character)
timeframe = f.readline()
timeframe = int(timeframe.rstrip('\n'))

#reads second line which is the Net ID (need to get rid of '\n' character)
login_id_noadd = f.readline()
login_id_noadd = login_id_noadd.rstrip('\n')

#reads third line which is the users login password (need to get rid of '\n' character)
login_pwd = f.readline()
login_pwd = login_pwd.rstrip('\n')

#close the file since we have all of the necessary information
f.close()

#set the login id for the logging into email
email_login_id = login_id_noadd + UCSB_ADD

#initialize our driver called web and use chrome to open up library booking link
web = webdriver.Chrome(HARDCODED_DRIVER_LOCATION)
web.get("http://libcal.library.ucsb.edu/rooms.php?i=12405")

#Note: The wait times were implemented merely to ensure that everything loaded before
#searching through the html file for data
web.implicitly_wait(0.5)

#get the current date from an xpath it is already in a better form to
#create a datetime object
current_date = web.find_element_by_xpath('//*[@id="s-lc-rm-tg-h"]')
current_date = current_date.get_attribute('innerHTML')
print "Today's day is currently: " + current_date + "\n"

#split the date and create a modified date in a format that will be used 
#to create a datetime object in the form Month day Year
mod_date = current_date.split(",",1)[1]
mod_date = mod_date.strip()
mod_date = mod_date.replace(',','')

#create the datetime object by passing mod_date into the string parse
#time function with the form Month day Year
mod_datetime = datetime.strptime(mod_date,'%B %d %Y')

#initialize the day of reference. This is important, because every number
#on the reservation grid is determined based on the numbers found on this
#day in the html file
day_reference = datetime(2017,12,30)

#the difference in days between the current day and the day of reference
#datetime objects will allow us to determine how many days it has been
#and thus how much to multiply each value by
numday_difference = mod_datetime - day_reference
numday_difference_str = str(numday_difference)
numday_difference_str = numday_difference_str.split(' ',1)[0]

#convert how many days it has been since the day of reference to an int
#for calculations later
numday_difference_int = int(numday_difference_str)

#dt is initialized to 14 days in advance (the soonest we can reserve a 
#room) and then it is added to the current day
dt = timedelta(days=FORTNIGHT)
mod_datetime = mod_datetime + dt


#these values are the initial values assigned to the grids in the html
#on Dec 30 2017 which is the day of reference. From here it is simply
#a matter of adding a multiple that was determined to be 816
if timeframe == 10:
	numdiff = 535731996
elif timeframe == 12:
	numdiff = 535732000
elif timeframe == 2:
	numdiff = 535732004
elif timeframe == 4:
	numdiff = 535732008
elif timeframe == 6:
	numdiff = 535732012
else:
	raise ValueError('a proper timeframe was not read from the file')
	
#perform basic arithmetic and cast as strings to later use when searching
#xpath to decide what grid boxes to click 
difference_factor = numdiff + (numday_difference_int+FORTNIGHT)*816
difference_factor1 = str(difference_factor)
difference_factor2 = str(difference_factor+1)
difference_factor3 = str(difference_factor+2)
difference_factor4 = str(difference_factor+3)

#go back to creating a datetime object of the day we are supposed to be 
#booking (14 days in advance)
mod_date = datetime.strftime(mod_datetime, '%b %d %Y')
print "Attempting to book a room on " + mod_date + "..\n"

#gather the month
month = mod_date.split(" ",3)[0]

#gather the day (removing any leading 0's because that does not follow
#the format that will later be used when searching xpath
day = mod_date.split(" ",3)[1]
day = day.lstrip('0')

#gather the year
year = mod_date.split(" ",3)[2]

#basic if statements checking so as not to reserve when unnecessary and
#because often days are unavailable for booking when school is not in 
#session 
if (month == 'Jan' and day < '16' and year == '2018'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')
elif (month == 'Feb' and day == '19' and year == '2018'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')
elif (month == 'Mar' and day > '22' and year == '2018'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')
elif (month == 'May' and day == '28' and year == '2018'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')
elif (month == 'Jun' and day > '14' and year == '2018'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')
elif (month == 'Jul' or month == 'Aug'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')
elif (month == 'Sep' and day < '24' and year == '2018'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')
elif (month == 'Nov' and (day == '12' or day == '22' or day == '23') and year == '2018'):
	raise ValueError(month + ' ' + day + ' ' + year + ' is not an academic day')

#click the month and then the day of the month that we will be booking
reserve_month = web.find_element_by_xpath("//*[text()[contains(., '"+ month +"')]]")
reserve_month.click()
time.sleep(0.5)
reserve_date = web.find_element_by_xpath("//*[@class='ui-state-default'][text()[contains(.,'"+ day +"')]]")
reserve_date.click()


#perform all of the bookings using our difference factors
#by clicking on the timegrids at predetermined locations
book1 = web.find_element_by_xpath("//*[@id='"+ difference_factor1 +"']")
book1.click()
time.sleep(0.3)
book2 = web.find_element_by_xpath("//*[@id='"+ difference_factor2 +"']")
book2.click()
time.sleep(0.3)
book3 = web.find_element_by_xpath("//*[@id='"+ difference_factor3 +"']")
book3.click()
time.sleep(0.3)
book4 = web.find_element_by_xpath("//*[@id='"+ difference_factor4 +"']")
book4.click()


#click continue
time.sleep(0.2)
cont = web.find_element_by_xpath('//*[@id="rm_tc_cont"]')
cont.click()

#click submit
time.sleep(0.5)
submit = web.find_element_by_xpath('//*[@id="s-lc-rm-sub"]')
submit.click()

#enter login credentials and click login
time.sleep(0.5)
username = web.find_element_by_id("username")
password = web.find_element_by_id("password")
username.send_keys(login_id_noadd)
password.send_keys(login_pwd)
log = web.find_element_by_xpath('//*[@id="fm1"]/section[3]/input[4]')
log.click()

#enter group name and submit the booking
time.sleep(0.5)
group_name = web.find_element_by_xpath('//*[@id="nick"]')
group_name.send_keys('CS nerds')
sub_booking = web.find_element_by_xpath('//*[@id="s-lc-rm-sub"]')
sub_booking.click()

#wait for the booking to send an email and then execute the following code
#after 5 minutes 
print "Waiting 5 minutes to ensure that confirmation email is sent and then accessing it..." 
time.sleep(300)

try:
	#open up a mail session with our IMTP_ADD and login then look at inbox
	mail = imaplib.IMAP4_SSL(IMTP_ADD)
	mail.login(email_login_id, login_pwd)
	mail.select('inbox')
	
	#search through all of the mail 
	type, data = mail.search(None, 'ALL')
	mail_ids = data[0]
	
	#make a list of the mail ids and assign the first (oldest) and latest
	#emails
	id_list = mail_ids.split()   
	first_email_id = int(id_list[0])
	latest_email_id = int(id_list[-1])

	#iterate through the most recent couple of emails to see if it came
	for i in range(latest_email_id,latest_email_id-2,-1):
		typ, data = mail.fetch(i,'(RFC822)')
		
		#iterate through the successful mail fetches
		for response_part in data:
			#if there is an instance find out what the subject and message
			#are. The confirmation email will be a multipart message 
			if isinstance(response_part, tuple):
				msg = email.message_from_string(response_part[1])
				email_subject = msg['subject']
				email_from = msg['from']
				
				#all confirmation emails follow this structure so if it
				#is a confirmation email gather the link and clean it up
				#so that it can be opened up and accessed.
				if email_subject == 'Please confirm your booking!' and email_from == 'LibCal <alerts@mail.libcal.com>':
					confirm_link = re.search("http://(.+?)\"", str(msg.get_payload(1)))
					unclean_link = str(confirm_link.groups(0))
					unclean_link = unclean_link.replace("'",'')
					unclean_link = unclean_link.replace("(",'')
					unclean_link = unclean_link.replace(",",'')
					unclean_link = unclean_link.replace(")",'')
					unclean_link = unclean_link.replace("amp;",'')
					clean_link = "http://" + unclean_link + "&m=confirm"
					print "accessing...\n%s\n" %(clean_link)
					urllib2.urlopen(str(clean_link))
					print "\n\nshould have worked\n\n"
							
except Exception as e:
	print str(e)
			
