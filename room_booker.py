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
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import smtplib
import re
import imaplib
import email
import logging
import sys

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# create a file handler
handler = logging.FileHandler('output.log')
handler.setLevel(logging.INFO)

# create a logging format
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(handler)

def getEmailDateTime(unformatted_date):
    return datetime.fromtimestamp(email.utils.mktime_tz(email.utils.parsedate_tz(unformatted_date)))

def main(NUM_OF_DAYS_IN_ADVANCE):
    global web
    global logger
    UCSB_ADD = "@umail.ucsb.edu"
    IMTP_ADD = "outlook.office365.com"
    IMTP_PORT = 993
    HARDCODED_DRIVER_LOCATION = 'C:\\Users\\Gavin\\Desktop\\chromedriver.exe'
    LIBCAL_EMAIL_ADDRESS = 'LibCal <alerts@mail.libcal.com>'
    driver_options = Options()
    driver_options.add_argument('--headless')
    driver_options.add_argument('--disable-gpu')

    # -- change this to the number of people whose credentials will be used --
    NUM_USERS = 6

    # -- set to True if you need to undo recent (today's) bookings for whatever reason -- WARNING: this will cancel ANY recent enough booking, which may be a booking you dont want cancelled
    RESET_BOOKINGS = False

    # -- Bookings older than this time window (in hours) will not be auto cancelled on bookings reset
    CANCEL_TIME_WINDOW = 1

    #input credentials
    #enters if file does not already exist and prompts user to enter necessary info to run program
    if (os.path.isfile('library_room.txt') == False):
        f = open('library_room.txt', 'w')
        for k in range(NUM_USERS):
        
            timeframe = input("Please enter\n11 to book 11-1\n1 to book 1-3\n 3 to book 3-5\n 5 to book 5-7\n or 7 to book 7-9\n")
            timeframe_check = input("Please enter the number again to confirm: ")
            while (timeframe != timeframe_check):
                print("Error the two numbers do not match. Please try again\n")
                timeframe = input("Please enter\n11 to book 11-1\n1 to book 1-3\n 3 to book 3-5\n 5 to book 5-7\n or 7 to book 7-9\n")
                timeframe_check = input("Please enter the number again to confirm: ")
            
            login_id_noadd = input("Please enter your Net ID (without the .umail extension): ")
            login_id_noadd_check = input("Please enter your Net ID again to confirm: ")
            while (login_id_noadd != login_id_noadd_check):
                print("Error the two Net ID's do not match. Please try again\n")
                login_id_noadd = input("Please enter your Net ID (without the umail extension): ")
                login_id_noadd_check = input("Please enter your Net ID again to confirm: ")
            
            login_pwd = input("Please enter your Net ID password: ")
            login_pwd_check = input("Please enter your password again to confirm: ")
            while (login_pwd != login_pwd_check):
                print("Error the two passwords do not match. Please try again\n")
                login_pwd = input("Please enter your Net ID password: ")
                login_pwd_check = input("Please enter your password again to confirm: ")
        
            #write the answers to the prompts line by line 
            f.write(timeframe + '\n')
            f.write(login_id_noadd + '\n')
            f.write(login_pwd + '\n')
            if (k == NUM_USERS):
                f.close()
                
    #opens file with necessary information		
    f = open('library_room.txt', 'r')

    #request timeslots 
    for k in range(NUM_USERS):
        #initialize our driver called web and use chrome to open up library booking link
        web = webdriver.Chrome(HARDCODED_DRIVER_LOCATION, chrome_options=driver_options)

        #reads first line which determines times to book (need to get rid of '\n' character)
        timeframe = f.readline()
        timeframe = int(timeframe.rstrip('\n'))

        #reads second line which is the Net ID (need to get rid of '\n' character)
        login_id_noadd = f.readline()
        login_id_noadd = login_id_noadd.rstrip('\n')

        #reads third line which is the users login password (need to get rid of '\n' character)
        login_pwd = f.readline()
        login_pwd = login_pwd.rstrip('\n')

        #email id
        email_login_id = login_id_noadd + UCSB_ADD

        logger.info("\n" + "-"*16 + "\n" + "Begin work on emailid: " + email_login_id)

        #cancelf previous bookings
        if (RESET_BOOKINGS):
            #open up a mail session with our IMTP_ADD and login then look at inbox
            mail = imaplib.IMAP4_SSL(IMTP_ADD)
            try:
                mail.login(email_login_id, login_pwd)
                mail.select('inbox')
                
                search_subject = 'Your booking has been confirmed!'

                #search through all of the mail 
                typ, data = mail.search(None, '(SUBJECT "' + search_subject + '")')
                mail_ids = data[0]
            
                #make a list of the mail ids and assign the latest email
                id_list = mail_ids.split()
                latest_email_id = int(id_list[-1])

                #check latest email to see if it came
                typ, data = mail.fetch(str(latest_email_id),'(RFC822)')
                msg = email.message_from_string(data[0][1].decode())
                email_subject = msg['subject']
                email_from = msg['from']
                email_time = getEmailDateTime(msg['Date'])
                curr_time = datetime.now()
                isRecent = (curr_time - email_time < timedelta(hours = CANCEL_TIME_WINDOW))
                
                #get booking confirmed email
                if email_subject == search_subject and email_from == LIBCAL_EMAIL_ADDRESS and isRecent:
                    cancel_link = re.search("http://(.+?)\"", str(msg.get_payload(0)))
                    unclean_link = str(cancel_link.groups(0))
                    unclean_link = unclean_link.replace("'",'')
                    unclean_link = unclean_link.replace("(",'')
                    unclean_link = unclean_link.replace(",",'')
                    unclean_link = unclean_link.replace(")",'')
                    unclean_link = unclean_link.replace("amp;",'')
                    clean_link = "http://" + unclean_link
                    logger.info("Accessing cancellation link for emailid: " + email_login_id +"\n\t->link = " + clean_link)
                    web.get(str(clean_link))
                    web.implicitly_wait(1)
                    try:
                        web.execute_script("document.getElementsByClassName('btn btn-primary')[0].click()")                        
                        logger.info("Booking should have cancelled")
                    except:
                        logger.warn("Stale cancellation link")
                else:
                    logger.warning("Could not cancel a booking for email id: " + email_login_id)
                                    
            except Exception as e:
                logger.warning('Failed to cancel email: ' + str(e), exc_info=True)
            mail.close()
            mail.logout()
        #get the library reservations page
        web.get("http://libcal.library.ucsb.edu/rooms.php?i=12405")

        #Note: The wait times were implemented merely to ensure that everything loaded before
        #searching through the html file for data
        web.implicitly_wait(0.5)

        #get the current date from an xpath it is already in a better form to
        #create a datetime object
        current_date = web.find_element_by_xpath('//*[@id="s-lc-rm-tg-h"]')
        current_date = current_date.get_attribute('innerHTML')
        logger.info("Today is " + current_date)

        #split the date and create a modified date in a format that will be used 
        #to create a datetime object in the form Month day Year
        mod_date = current_date.split(",",1)[1]
        mod_date = mod_date.strip()
        mod_date = mod_date.replace(',','')

        #create the datetime object by passing mod_date into the string parse
        #time function with the form Month day Year
        current_datetime = datetime.strptime(mod_date,'%B %d %Y')

        #initialize the day of reference. This is important, because every number
        #on the reservation grid is determined based on the numbers found on this
        #day in the html file
        day_reference = datetime(2018,1,17)

        #dt is initialized to 14 days in advance (the soonest we can reserve a 
        #room) and then it is added to the current day
        dt = timedelta(days=NUM_OF_DAYS_IN_ADVANCE)
        booking_datetime = current_datetime + dt

        #the difference in days between the current day and the day of reference
        #datetime objects will allow us to determine how many days it has been
        #and thus how much to multiply each value by
        numday_difference = booking_datetime - day_reference

        #these values are the initial values assigned to the grids in the html
        #on Dec 30 2017 which is the day of reference. From here it is simply
        #a matter of adding a multiple that was determined to be 816
        if timeframe == 11:
            numdiff = 600893789
        elif timeframe == 1:
            numdiff = 600893793
        elif timeframe == 3:
            numdiff = 600893797
        elif timeframe == 5:
            numdiff = 600893801
        elif timeframe == 7:
            numdiff = 600893805
        elif timeframe == 9:
            numdiff = 600893809
        else:
            logger.warning('a proper timeframe was not read from the file. Timeframe was: ' + str(timeframe))
            web.close()
            continue
        
        #perform basic arithmetic and cast as strings to later use when searching
        #xpath to decide what grid boxes to click 
        difference_factor = numdiff + (numday_difference.days)*816
        difference_factor1 = str(difference_factor)
        difference_factor2 = str(difference_factor+1)
        difference_factor3 = str(difference_factor+2)
        difference_factor4 = str(difference_factor+3)

        #go back to creating a datetime object of the day we are supposed to be 
        #booking (14 days in advance)
        mod_date = datetime.strftime(booking_datetime, '%b %d %Y')
        logger.info("Attempting to book a room on " + mod_date + " (" + str(days_in_advance) + " days from today) for emailid: " + email_login_id)

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
        if ((month == 'Jan' and day < '16' and year == '2018')
        or (month == 'Feb' and day == '19' and year == '2018')
        or (month == 'Mar' and day > '22' and year == '2018')
        or (month == 'May' and day == '28' and year == '2018')
        or (month == 'Jun' and day > '14' and year == '2018')
        or (month == 'Jul' or month == 'Aug')
        or (month == 'Sep' and day < '24' and year == '2018')
        or (month == 'Nov' and (day == '12' or day == '22' or day == '23') and year == '2018')):
            logger.error(month + ' ' + day + ' ' + year + ' is not an academic day')
            web.close()
            sys.exit()

        #prepare a wait to ensure everything is clicked
        wait = WebDriverWait(web, 10)

        #click the month and then the day of the month that we will be booking
        reserve_month = web.find_element_by_xpath("//*[text()[contains(., '"+ month +"')]]")
        reserve_month.click()
        time.sleep(0.5)
        reserve_date = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@class='ui-state-default'][text()[contains(.,'"+ day +"')]]")))
        reserve_date.click()
        time.sleep(1)

        #perform all of the bookings using our difference factors
        #by clicking on the timegrids at predetermined locations
        try:
            web.execute_script("document.getElementById('" + difference_factor1 + "').click()")
            web.execute_script("document.getElementById('" + difference_factor2 + "').click()")
            web.execute_script("document.getElementById('" + difference_factor3 + "').click()")
            web.execute_script("document.getElementById('" + difference_factor4 + "').click()")
        except:
            logger.warn("Time slots unavailable for: " + str(timeframe) + " - " + str((timeframe+2) % 12) + " on " + month + "-" + day + "-" + year)
            web.close()
            continue #continue loop for next credential

        #click continue
        cont = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rm_tc_cont"]')))
        cont.click()

        #click submit
        submit = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="s-lc-rm-sub"]')))
        submit.click()
        time.sleep(0.7)

        #enter login credentials and click login
        username = wait.until(EC.presence_of_element_located((By.ID, "username")))
        password = web.find_element_by_id("password")
        username.send_keys(login_id_noadd)
        password.send_keys(login_pwd)
        log = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fm1"]/section[3]/input[4]')))
        log.click()

        #enter group name and submit the booking
        group_name = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="nick"]')))
        group_name.send_keys('CS nerds')
        sub_booking = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="s-lc-rm-sub"]')))
        time.sleep(0.7)
        sub_booking.click()
        web.implicitly_wait(1)

        #wait for the booking to send an email and then execute the following code
        #after 5 minutes 
        logger.info("Waiting to ensure that confirmation email is sent and then accessing it...")
        time.sleep(10)

        #open up a mail session with our IMTP_ADD and login then look at inbox
        mail = imaplib.IMAP4_SSL(IMTP_ADD)

        #confirm emails
        try:
            mail.login(email_login_id, login_pwd)
            mail.select('inbox')

            search_subject = 'Please confirm your booking!'

            #search through all of the mail 
            typ, data = mail.search(None, '(SUBJECT "' + search_subject + '")')
            mail_ids = data[0]
        
            #make a list of the mail ids and assign the first (oldest) and latest
            #emails
            id_list = mail_ids.split()
            latest_email_id = int(id_list[-1])

            #check latest email to see if it came
            typ, data = mail.fetch(str(latest_email_id),'(RFC822)')
            
            #The confirmation email will be a multipart message 
            msg = email.message_from_string(data[0][1].decode())
            email_subject = msg['subject']
            email_from = msg['from']

            email_time = getEmailDateTime(msg['Date'])
            curr_time = datetime.now()
            isRecent = curr_time - email_time < timedelta(hours = CANCEL_TIME_WINDOW)

            #all confirmation emails follow this structure so if it
            #is a confirmation email gather the link and clean it up
            #so that it can be opened up and accessed.
            if email_subject == search_subject and email_from == 'LibCal <alerts@mail.libcal.com>' and isRecent:
                confirm_link = re.search("http://(.+?)\"", str(msg.get_payload(1)))
                unclean_link = str(confirm_link.groups(0))
                unclean_link = unclean_link.replace("'",'')
                unclean_link = unclean_link.replace("(",'')
                unclean_link = unclean_link.replace(",",'')
                unclean_link = unclean_link.replace(")",'')
                unclean_link = unclean_link.replace("amp;",'')
                clean_link = "http://" + unclean_link + "&m=confirm"
                logger.info("Accessing confirmation link for emailid: " + email_login_id +"\n\t->link = " + clean_link)
                web.get(str(clean_link))
                web.implicitly_wait(1)
                logger.info("Booking should have confirmed")
            else:
                logger.warning("Could not confirm booking for timeslot: " + str(timeframe) + "-" + str((timeframe+2)%12))
        except Exception as e:
            logger.warning('Failed to confirm email: ' + str(e), exc_info=True)
        mail.close()
        mail.logout()
        web.close()
    f.close()

try:
    lower_bound = 14 #inclusive
    upper_bound = 14 #inclusive
    for days_in_advance in range(lower_bound, upper_bound + 1):
        main(days_in_advance)
except Exception as e:
    web.close()
    logger.error("Something happened: " + str(e), exc_info=True)