#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  room_booker.py
#  Original Author: rodericd5
#  Contributer: GavinFrazar
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
handler = logging.FileHandler('output' + sys.argv[1] + '.log')
handler.setLevel(logging.INFO)

# create a logging format
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(handler)

def getEmailDateTime(unformatted_date):
    return datetime.fromtimestamp(email.utils.mktime_tz(email.utils.parsedate_tz(unformatted_date)))

def run(NUM_OF_DAYS_IN_ADVANCE=14, STARTING_TIMESLOT=11, USERS=None, RESET_BOOKINGS=False, CANCEL_TIME_WINDOW=1, HEADLESS=False, HARDCODED_DRIVER_LOCATION="chromedriver", ROOM_OFFSET=0):
    global web
    global logger
    UCSB_ADD = "@ucsb.edu"
    IMTP_ADD = "imap.gmail.com"
    #IMTP_PORT = 993
    LIBCAL_EMAIL_ADDRESS = 'LibCal <alerts@mail.libcal.com>'
    successful_users = set()

    #request timeslots 
    for k, key in enumerate(USERS):
        #initialize our driver called web and use chrome to open up library booking link
        
        if (HEADLESS):
            driver_options = Options()
            driver_options.add_argument('--headless')
            driver_options.add_argument('--disable-gpu')
            web = webdriver.Chrome(HARDCODED_DRIVER_LOCATION, chrome_options=driver_options)
        else:
            web = webdriver.Chrome(HARDCODED_DRIVER_LOCATION)

        #reads first line which is the Net ID (need to get rid of '\n' character)
        login_id_noadd = key

        #reads second line which is the users login password (need to get rid of '\n' character)
        login_pwd = USERS[key]

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
                    cancel_link = re.search("https://(.+?)\"", str(msg.get_payload(0)))
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
        
        while (True):
            #get the library reservations page
            web.get("http://libcal.library.ucsb.edu/rooms.php?i=12405")

            #Note: The wait times were implemented merely to ensure that everything loaded before
            #searching through the html file for data
            web.implicitly_wait(0.5)

            # Get the actual current date
            local_datetime = datetime.now()

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
            current_server_datetime = datetime.strptime(mod_date,'%B %d %Y')
            if (local_datetime - current_server_datetime).days == 0:
                break
            else:
                logger.warning("Server time out of sync with local time. Retrying...")

        #initialize the day of reference. This is important, because every number
        #on the reservation grid is determined based on the numbers found on this
        #day in the html file
        day_reference = datetime(2018,9,28)
        timeslot_on_day_reference = 647733591

        #dt is initialized to 14 days in advance (the soonest we can reserve a 
        #room) and then it is added to the current day
        dt = timedelta(days=NUM_OF_DAYS_IN_ADVANCE)
        booking_datetime = current_server_datetime + dt

        #the difference in days between the current day and the day of reference
        #datetime objects will allow us to determine how many days it has been
        #and thus how much to multiply each value by
        numday_difference = booking_datetime - day_reference
        logger.info('Booking date is ' + str(numday_difference) + ' from day of reference (' + str(day_reference) +')')
        #the beginning timeslot of the 2 hours we want to book
        TARGET_TIMESLOT = STARTING_TIMESLOT + 2*k

        #this value is the initial value assigned to room 2528 grid in the html
        #on Jan 01 2018 for 11am. From here it is simply
        #a matter of adding a multiple of 816 per day since that day.
        REFERENCE_TIMESLOT = 11 # our reference timeslot was for 11am
        REFERENCE_TIMESLOT_ID = timeslot_on_day_reference + 2*(TARGET_TIMESLOT - REFERENCE_TIMESLOT) + ROOM_OFFSET
        
        #perform basic arithmetic and cast as strings to later use when searching
        #xpath to decide what grid boxes to click 
        difference_factor = REFERENCE_TIMESLOT_ID + (numday_difference.days)*816
        logger.info("Timeslot grid id was: " + str(difference_factor))

        #timeslots are reserved 30 minutes at a time and for a maximum of 2 hours, so we need to reserve four slots
        difference_factor1 = str(difference_factor)
        difference_factor2 = str(difference_factor+1)
        difference_factor3 = str(difference_factor+2)
        difference_factor4 = str(difference_factor+3)

        #Used for logging purposes
        TIMESLOT_RANGE = str((TARGET_TIMESLOT) % 12) + " - " + str((TARGET_TIMESLOT + 2) % 12)

        #go back to creating a datetime object of the day we are supposed to be 
        #booking (14 days in advance)
        mod_date = datetime.strftime(booking_datetime, '%b %d %Y')
        logger.info("Attempting to book the room for " + TIMESLOT_RANGE + " on " + mod_date + " (" + str(NUM_OF_DAYS_IN_ADVANCE) + " days from today) for emailid: " + email_login_id)

        #gather the month
        month = mod_date.split(" ",3)[0]

        #gather the day (removing any leading 0's because that does not follow
        #the format that will later be used when searching xpath
        day = mod_date.split(" ",3)[1]
        day = day.lstrip('0')

        #gather the year
        year = mod_date.split(" ",3)[2]

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
            logger.warn("Time slots unavailable for: " + TIMESLOT_RANGE + " on " + month + "-" + day + "-" + year)
            web.quit()
            continue #continue loop for next credential

        #click continue
        cont = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="rm_tc_cont"]')))
        cont.click()

        #click submit
        submit = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="s-lc-rm-sub"]')))
        submit.click()
        time.sleep(0.5)

        #enter login credentials and click login
        username = wait.until(EC.presence_of_element_located((By.ID, "username")))
        password = web.find_element_by_id("password")
        username.send_keys(login_id_noadd)
        password.send_keys(login_pwd)
        login = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fm1"]/section[3]/input[4]')))
        login.click()

        #enter group name and submit the booking
        group_name = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="nick"]')))
        group_name.send_keys('CS nerds')
        sub_booking = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="s-lc-rm-sub"]')))
        sub_booking.click()
        web.implicitly_wait(1)
        web.quit()
        successful_users.add(key)

    for key in successful_users:
        if (HEADLESS):
            driver_options = Options()
            driver_options.add_argument('--headless')
            driver_options.add_argument('--disable-gpu')
            web = webdriver.Chrome(HARDCODED_DRIVER_LOCATION, chrome_options=driver_options)
        else:
            web = webdriver.Chrome(HARDCODED_DRIVER_LOCATION)
        #wait for the booking to send an email and then execute the following code
        #after 5 minutes 
        logger.info("Waiting to ensure that confirmation email is sent and then accessing it...")
        time.sleep(10)  
        #open up a mail session with our IMTP_ADD and login then look at inbox
        mail = imaplib.IMAP4_SSL(IMTP_ADD)

        email_login_id = key + UCSB_ADD
        #confirm emails
        try:
            # search at most twice for a confirmation email
            for _ in range(2):
                try:
                    mail.login(email_login_id, USERS[key])
                    mail.select('inbox')

                    search_subject = 'Please confirm your booking!'

                    #search through all of the mail 
                    typ, data = mail.search(None, '(SUBJECT "' + search_subject + '")')
                    mail_ids = data[0]
                
                    #make a list of the mail ids and assign the first (oldest) and latest
                    #emails
                    id_list = mail_ids.split()
                    latest_email_id = int(id_list[-1])
                except:
                    mail.close()
                    mail.logout()
                    logger.warning("Could not find confirm email. Waiting before retry")
                    time.sleep(20)
                    logger.info('Retrying search for confirm email')
                else:
                    break

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
            if email_subject == search_subject and email_from == LIBCAL_EMAIL_ADDRESS and isRecent:
                # logger.info('payload: ' + str(msg.get_payload(1)))
                confirm_link = re.search("https://(.+?)\"", str(msg.get_payload(1)))
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
                logger.warning("Could not confirm booking for timeslots: " + TIMESLOT_RANGE)
        except Exception as e:
            logger.warning('Failed to confirm email: ' + str(e), exc_info=True)
        mail.close()
        mail.logout()
        web.quit()

def main(users, starting_timeslot):    
    lower_bound = 14 #inclusive
    upper_bound = 14 #inclusive

    # -- set to True if you need to undo recent (today's) bookings for whatever reason -- WARNING: this will cancel ANY recent enough booking, which may be a booking you dont want cancelled
    RESET_BOOKINGS = True

    # -- Bookings older than this time window (in hours) will not be auto cancelled on bookings reset
    CANCEL_TIME_WINDOW = 1

    # -- Boolean, if true the webdriver will be run in headless mode
    HEADLESS = True
    
    # -- you can leave this as it is if you set up your path variables to point to your chromedriver or if you simply put the chromedriver in the same folder as this py script
    # -- Otherwise set it to 'path/to/your/chromedriver.exe'
    HARDCODED_DRIVER_LOCATION = "chromedriver"

    room_offset = 0

    global web
    web = None
    try:
        for days_in_advance in range(lower_bound, upper_bound + 1):
            run(days_in_advance, starting_timeslot, users, RESET_BOOKINGS, CANCEL_TIME_WINDOW, HEADLESS, HARDCODED_DRIVER_LOCATION, room_offset)
    except Exception as e:
        logger.error("Something happened: " + str(e), exc_info=True)
    finally:
        if (web):
            web.quit()

if __name__ == '__main__':
    import multiprocessing
    import login_info

    user_number = int(sys.argv[1])
    if user_number < 0 or user_number >= len(login_info.users):
        logger.warning("Invalid user index. Got: " + str(user_number))
    
    main(login_info.users[user_number], 11 + 2*user_number)