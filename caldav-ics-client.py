#!/usr/bin/python

#
## caldav-ics-client.py
# Receives event details from arguments, parse them, create an ICS event, upload it via webdav PUT request.
# Must provide JSON formatted file with: User-Agent, webdav server url, authentication, etc.
# Based on a NextCloud environment.
#
# See README.me for full details.
#
## License
# Released under GPL-3.0 license.
#
# 2023.06.21
# https://github.com/ynad/caldav-py-handler
# info@danielevercelli.it
#

###################################################################################################
# APP SETTINGS - do not edit
version_num="0.4.9"
update_version_url="https://raw.githubusercontent.com/ynad/caldav-py-handler/main/VERSION"
update_url="https://raw.githubusercontent.com/ynad/caldav-py-handler/main/caldav-ics-client.py"
user_agent=f"caldav-ics-client/{version_num}"
ics_file="tmp_caldav-ics-event.ics"
logging_file="debug.log"
###################################################################################################



import sys, os, logging
import json
import click
import requests
import urllib.request
import random
import signal
from typing import Dict, List, Tuple
from requests.auth import HTTPBasicAuth
from icalendar import Calendar, Event, Alarm, vCalAddress, vText
from datetime import datetime, timedelta
from pathlib import Path



# Enable logging
logging_file = f"{os.path.dirname(__file__)}/{logging_file}"
logging.basicConfig(
    filename=logging_file,
    filemode='w',
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)
logger = logging.getLogger(__name__)



# show command syntax
def print_header(user_settings) -> None:
    print(f"\nCaldav ICS CLIent - v{version_num} - {user_settings['domain']}\n"
           "==============================================\n")


# show command syntax
def show_syntax() -> None:
    print(f"\nMissing or wrong arguments! Syntax:\n\n{sys.argv}\n\n"
            "    --name \"event name\"\n"
            "    --descr \"event description\"\n"
            "    --start_day dd/mm/YYYY [dd/mm/YYYY [...]]\n"
            "    --end_day dd/mm/YYYY [dd/mm/YYYY [...]]\n"
            "   [--start_hr HH:MM [HH:MM [...]]]\n"
            "   [--end_hr HH:MM [HH:MM [...]]]\n"
            "   [--loc \"event location\"]\n"
            "   [--cal \"calendar to be used\"]\n"
            "   [--invite : email(s) to be invited, separated by space]\n"
            "   [--alarm_type : alarm to be set on event: \"DISPLAY\" or \"EMAIL\". Default: none]\n"
            "   [--alarm_format : \"h\" = hours, \"d\" = days]\n"
            "   [--alarm_time : time before the event to set an alarm for]\n"
            "   [--prompt \"y/n\" : wait or skip user confirmation. Default: y]\n"
            "   [--config \"path\\to\\config-file.json\"]\n"
    )


# load user settings from file
def load_user_settings(user_config) -> dict:
    if os.path.exists(user_config):
        logger.info("Loading user settings JSON from file: " + str(user_config))
        with open(user_config, 'r') as f:
            return json.load(f)
    else:
        logger.error(f"User settings JSON missing: {user_config}")
        return None


# check software updates
def check_updates() -> None:

    # get latest version number
    response = requests.get(update_version_url)
    if (response.status_code == 200):
        update_version = response.text.split('\n')

        # newer version available on repo
        if update_version[0] > version_num:
            logger.debug(f"Current version: {version_num}, found update: {update_version}")
            print(f"A new version is available: {update_version[0]}, {update_version[1]}\n"
                   "After the update you may have to re-launch the program.\n"
                   "Do you want to update now? (Y/N)")
            run_update = input()

            if run_update.upper() == 'Y':
                logger.debug("Downloading new version and restarting")
                print(f"Downloading new version and restarting...")
                urllib.request.urlretrieve(update_url, __file__)
                os.execl(sys.executable, 'python', __file__, *sys.argv[1:])
            else:
                logger.debug("Update skipped")
                print("Update skipped.")
    else:
        e = "Error occurred while checking available updates."
        logger.error(f"{e} {response.status_code}, {response.text}")
        print(e)


# check string date format
def check_date(date) -> Tuple[bool, str]:
    try:
        datetime.strptime(date, "%d/%m/%Y")
    except ValueError as err:
        return False, err
    return True, ""


# check string time format
def check_time(time) -> Tuple[bool, str]:
    try:
        datetime.strptime(time, "%H:%M")
    except ValueError as err:
        return False, err
    return True, ""


# check if first date is after the second one
def is_after_date(date_i, date_j) -> bool:
    return datetime.strptime(date_i, "%d/%m/%Y") > datetime.strptime(date_j, "%d/%m/%Y")


# check if first hour is after the second one
def is_after_hour(date_i, date_j) -> bool:
    return datetime.strptime(date_i, "%H:%M") > datetime.strptime(date_j, "%H:%M")


# check arguments and return error strings
def args_check(user_settings, start_day, end_day, start_hr, end_hr) -> Tuple[bool, str]:

    logger.debug(f"Running args check")

    # start and end day are mandatory
    if not start_day or not end_day:
        return False, "Event start and end date are mandatory!"

    # check START date format
    start_day_list = start_day.split()
    for start_d in start_day_list:
        date_ok, date_err = check_date(start_d)
        if not date_ok:
            return False, date_err

    # check END date format
    end_day_list = end_day.split()
    for end_d in end_day_list:
        date_ok, date_err = check_date(end_d)
        if not date_ok:
            return False, date_err

    # check list lenght, must be equal for start and end days
    if len(start_day_list) != len(end_day_list):
        return False, "Start and end days count cannot differ!"

    # check START date is not after END date
    for i, day in enumerate(start_day_list):
        if is_after_date(start_day_list[i], end_day_list[i]):
            err = f"Event start date cannot be after end date: {start_day_list[i]}, {end_day_list[i]}"
            return False, err

    # check time format, if any is given
    if start_hr and end_hr:

        start_hr_list = start_hr.split()
        for start_hour in start_hr_list:
            time_ok, time_err = check_time(start_hour)
            if not time_ok:
                return False, time_err

        end_hr_list = end_hr.split()
        for end_hour in end_hr_list:
            time_ok, time_err = check_time(end_hour)
            if not time_ok:
                return False, time_err

        # check list lenght, must be equal for start and end hours as well as for days count
        if (len(start_hr_list) != len(end_hr_list)) or (len(start_hr_list) != len(start_day_list)):
            return False, "Start and end hours count cannot differ!"

        # check START hour is not after END hour
        for i, hour in enumerate(start_hr_list):
            if is_after_hour(start_hr_list[i], end_hr_list[i]):
                err = f"Event start hour cannot be after end hour: {start_hr_list[i]}, {end_hr_list[i]}"
                return False, err

    return True, ""


# create ICS file with provided event details
def create_ics(user_settings, event_details) -> None:

    # init calendar
    logger.debug(f"ICS: create calendar")
    mycal = Calendar()
    # set properties to be compliant
    mycal.add("prodid", f"-//CalDav ICS Client//{version_num}//{user_settings['domain']}//github.com/ynad/caldav-py-handler//")
    mycal.add("version", "2.0")
    #mycal.add('method', "REQUEST")

    # add calendar subcomponents
    logger.debug(f"ICS: create event")
    myevent = Event()
    #myevent.add('name', event_details['name'])
    myevent.add('summary', event_details['name'])
    myevent.add('description', event_details['description'])
    myevent.add('dtstart', event_details['start'])
    myevent.add('dtend', event_details['end'])
    myevent.add('status', "confirmed")

    # add location
    myevent['location'] = vText(event_details['location'])

    # creation time
    create_time = datetime.strptime(datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "%d/%m/%Y %H:%M:%S")
    logger.debug(f"ICS: created time: {create_time}")
    myevent.add('created', create_time)
    
    # uid - unique event ID
    myevent['uid'] = event_details['uid']
    myevent.add('priority', 5)

    # add organizer
    logger.debug(f"ICS: organizer: {user_settings['organizer_email']}")
    organizer = vCalAddress(f"MAILTO:{user_settings['organizer_email']}")
    organizer.params['name'] = vText(user_settings['organizer_name'])
    organizer.params['role'] = vText(user_settings['organizer_role'])
    myevent['organizer'] = organizer

    # add invites if present
    if 'invite' in event_details:
        invitees = event_details['invite'].split()
        for i in invitees:
            logger.debug(f"ICS: adding invite for: {i}")
            attendee = vCalAddress(f"MAILTO:{i}")
            attendee.params['name'] = vText(i)
            attendee.params['role'] = vText('REQ-PARTICIPANT')
            myevent.add('attendee', attendee, encode=0)

    # add an alarm for the event
    if 'alarm_type' in event_details:
        logger.debug(f"ICS: adding alarm")
        myalarm = Alarm()
        myalarm.add("action", event_details['alarm_type'])
        myalarm.add('summary', event_details['name'])
        myalarm.add('description', event_details['description'])

        # if invitees are present, add email notification
        if 'invite' in event_details:
            invitees = event_details['invite'].split()
            for i in invitees:
                logger.debug(f"ICS: alarm invite for: {i}")
                attendee = vCalAddress(f"MAILTO:{i}")
                #attendee.params['name'] = vText(i)
                #attendee.params['role'] = vText('REQ-PARTICIPANT')
                myalarm.add('attendee', attendee, encode=0)

        # set trigger time
        #myalarm.add("trigger", timedelta(days=-reminder_days))
        if event_details['fullday'] is True:
            logger.debug(f"ICS: fullday event")
            myalarm.add("TRIGGER;RELATED=START", f"-P{event_details['alarm_time']}{event_details['alarm_format']}")
        else:
            logger.debug(f"ICS: fixed hours event")
            myalarm.add("TRIGGER;RELATED=START", f"-PT{event_details['alarm_time']}{event_details['alarm_format']}")
        myevent.add_component(myalarm)

    # add event to the calendar
    mycal.add_component(myevent)

    # write event to ICS file
    logger.info(f"ICS: write to file {ics_file}")
    with open(ics_file, 'wb') as f:
        f.write(mycal.to_ical())


# make PUT request to upload ICS event file to given calendar
def webdav_put_ics(user_settings, calendar, event_id) -> None:

    # check ICS existance
    if not os.path.exists(ics_file):
        err = f"ERROR: missing ICS file {ics_file}, can't continue"
        print(err)
        logger.error(err)
        return

    # read ICS ifle
    logger.info(f"webdav: read ics from file {ics_file}")
    with open(ics_file, 'rb') as f:
        data = f.read()

    # if calendar is not set go default
    if calendar == None:
        calendar = user_settings['calendar_default']
    logger.debug(f"webdav: calendar set to: {calendar}")

    # make PUT request
    try:
        logger.debug(f"webdav: put request user-agent: {user_agent}")
        res = requests.put(url=f"{user_settings['server']}/{user_settings['username']}/{calendar}/{event_id}",
                            data=data,
                            headers={'Content-Type': 'text/calendar', 'User-Agent': user_agent},
                            auth = HTTPBasicAuth(user_settings['username'], user_settings['password']))

        if (res.status_code == 201):
            msg = f"Created ({res.status_code})"
            print(msg)
            logger.info(msg)

        elif (res.status_code == 202):
            msg = f"Accepted ({res.status_code})"
            print(msg)
            logger.info(msg)

        elif (res.status_code == 204):
            msg = f"No Content ({res.status_code})"
            print(msg)
            logger.info(msg)

        else:
            err = (
                f"ERROR: {str(res.status_code)}, {res.text}"
                + "; in response from server."
            )
            raise Exception(err)

    except Exception as e:
        print(e)
        logger.error(e)

    finally:
        # rm tmp ics files
        if os.path.exists(ics_file):
            logger.info(f"webdav: remove ics file {ics_file}")
            os.remove(ics_file)
        #pass



@click.command()
@click.option(
    "--config",
    type=str,
    default="user_settings.json"
)
@click.option(
    "--name",
    type=str,
    default="default event title"
)
@click.option(
    "--descr",
    type=str,
    default="default event description"
)
@click.option(
    "--start_day",
    type=str,
    default=""
)
@click.option(
    "--start_hr",
    type=str,
    default=""
)
@click.option(
    "--end_day",
    type=str,
    default=""
)
@click.option(
    "--end_hr",
    type=str,
    default=""
)
@click.option(
    "--loc",
    type=str,
    default=""
)
@click.option(
    "--cal",
    type=str,
    default="personal"
)
@click.option(
    "--invite",
    type=str,
    default=""
)
@click.option(
    "--alarm_type",
    type=str,
    default=""
)
@click.option(
    "--alarm_format",
    type=str,
    default=""
)
@click.option(
    "--alarm_time",
    type=str,
    default=""
)
@click.option(
    "--prompt",
    type=str,
    default="y"
)



## Main
def main(config, name, descr, start_day, start_hr, end_day, end_hr, loc, cal, invite, alarm_type, alarm_format, alarm_time, prompt):

    # load user settings from json file
    user_settings = load_user_settings(config)
    if not user_settings:
        print(f"User config file missing: {config}")
        input("Cannot continue. Press enter to exit.")
        return

    # print software header
    print_header(user_settings)

    # check software updates
    check_updates()

    # check command line arguments
    args_ack, err = args_check(user_settings, start_day, end_day, start_hr, end_hr)
    if not args_ack:
        logger.error(err)
        show_syntax()
        print(f"{err}\n")
        input("Press enter to exit.")
        return

    # if more than one start & end days/hours are provided, split them and repeat all procedure for each one
    start_day_list = start_day.split()
    end_day_list = end_day.split()

    if start_hr and end_hr:
        start_hr_list = start_hr.split()
        end_hr_list = end_hr.split()
    
    events_list = []

    # cycle by key over list of event dates and to list one event each
    for i, day in enumerate(start_day_list):

        # build event details
        event_details = {
            'name' : name,
            'description' : descr,
            'calendar' : cal if cal else None,
            'location' : loc if loc else user_settings['location_default'],
            'uid' : (f"{str(datetime.now().timestamp())}_{random.randint(100000, 999999)}_{name}@{user_settings['domain']}").replace(" ", "-")
        }
        logger.debug(f"Building event details with UID: {event_details['uid']}")

        # event with fixed hours
        if start_hr and end_hr:
            # hours set to 00:00 equals full day event
            if (start_hr_list[i] == "00:00") and (end_hr_list[i] == "00:00"):
                event_details.update( { 'start' : datetime.strptime(f"{start_day_list[i]}", "%d/%m/%Y").date() } )
                event_details.update( { 'end' : datetime.strptime(f"{end_day_list[i]}", "%d/%m/%Y").date() + timedelta(days=1) } )
                event_details.update( { 'fullday' : True } )
                logger.debug(f"Full day event, all-0 hours")
            else:
            # set fixed hours
                event_details.update( { 'start' : datetime.strptime(f"{start_day_list[i]} {start_hr_list[i]}", "%d/%m/%Y %H:%M") } )
                event_details.update( { 'end' : datetime.strptime(f"{end_day_list[i]} {end_hr_list[i]}", "%d/%m/%Y %H:%M") } )
                event_details.update( { 'fullday' : False } )
                logger.debug(f"Fixed hours event")
        # full day event
        else:
            event_details.update( { 'start' : datetime.strptime(f"{start_day_list[i]}", "%d/%m/%Y").date() } )
            event_details.update( { 'end' : datetime.strptime(f"{end_day_list[i]}", "%d/%m/%Y").date() + timedelta(days=1) } )
            event_details.update( { 'fullday' : True } )
            logger.debug(f"Full day event")
        
        # add invitees, can be 1 or more separated by a space
        if invite:
            event_details.update( { 'invite' : invite } )
            logger.debug(f"Invites requested for: {invite}")

        # set alarm - all 3 parameters must be given otherwise none is set
        if alarm_type and alarm_format and alarm_time:
            alarm_type = alarm_type.upper()
            alarm_format = alarm_format.upper()
            if ( alarm_type == 'DISPLAY' or alarm_type == 'EMAIL') and ( alarm_format == 'H' or alarm_format == 'D' ):
                event_details.update( { 'alarm_type' : alarm_type } )
                event_details.update( { 'alarm_format' : alarm_format } )
                event_details.update( { 'alarm_time' : alarm_time } )
                logger.debug(f"Alarm requested: {alarm_type}, {alarm_format}, {alarm_time}")

        # append event to list
        events_list.append(event_details)


    # wait for user confirmation if enabled. To skip give argument '--prompt n'
    if prompt == "y":
        logger.debug(f"Wait for user prompt to proceed")
        print(f"\nThe following {len(start_day_list)} event(s) will be added:\n")

        # cycle over events list
        for j, event_n in enumerate(events_list):

            print(f"Event {j+1}/{len(start_day_list)}\n"
                  f"-----------\n"
                  f"NAME:\t\t{event_n['name']}\n"
                  f"DESCRIPTION:\t{event_n['description']}\n"
                  f"\nSTART DATE:\t{datetime.strftime(event_n['start'], '%d/%m/%Y %H:%M:%S')}\n"
                  f"END DATE:\t{datetime.strftime(event_n['end'], '%d/%m/%Y %H:%M:%S')}\n"
                  f"LOCATION:\t{event_n['location']}\n"
                  f"CALENDAR:\t{event_n['calendar']}")
            if invite:
                print(f"INVITEE:\t{event_n['invite']}")
            if 'alarm_type' in event_n:
                print(f"ALARM:\t\t{event_n['alarm_type']}, {event_n['alarm_time']}{event_n['alarm_format']} before")
            print("----------------------------------------------\n")

        input("Press enter to confirm.")

    # compile ICS and create event for each one in list
    for event_n in events_list:

        # compile ICS file
        create_ics(user_settings, event_n)

        # upload it to caldav server
        webdav_put_ics(user_settings, event_n['calendar'], event_n['uid'])

    # final prompt
    if prompt == "y":
        print("")
        input("Press enter to exit.")


if __name__ == '__main__':
    main()
