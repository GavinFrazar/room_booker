#!/bin/bash
for i in {0..5}
do
    python room_booker.py $i &
done

#options:
#--from <days from now to begin booking range>
#--to <days from now to end booking range>
#--reset makes recent bookings reset
#--headless <true/false>
#--earliest-time <time of day>
#--room <room number to book>