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
# v0.4.2 - 2023.02.08 - https://github.com/ynad/caldav-py-handler
# info@danielevercelli.it
#

###################################################################################################
# USER SETTINGS - adjust to your environment
user_conf_json="user_settings.json"
logging_file="debug.log"
prompt_wait=True

# APP SETTINGS - no need to edit normally
version_num="0.4.3"
user_agent=f"caldav-ics-client-{version_num}"
ics_file="tmp_caldav-ics-event.ics"
###################################################################################################



import sys, os, logging
import json
import click
import requests
from typing import Dict, List, Tuple
from requests.auth import HTTPBasicAuth
from icalendar import Calendar, Event, Alarm, vCalAddress, vText
from datetime import datetime, timedelta
from pathlib import Path



# Enable logging
logging.basicConfig(
    filename=logging_file,
    filemode='w',
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)
logger = logging.getLogger(__name__)



# show command syntax
def show_syntax(user_settings) -> str:
    err = (
        f"\nCaldav ICS CLIent - v{version_num} - {user_settings['domain']}\n"
        "==============================================\n\n"
        f"Missing or wrong arguments! Syntax:\n{sys.argv}\n"
        "    --name \"event name\"\n"
        "    --descr \"event description\"\n"
        "    --start_day dd/mm/YYYY\n"
        "    --end_day dd/mm/YYYY\n"
        "   [--start_hr HH:MM:SS]\n"
        "   [--end_hr HH:MM:SS]\n"
        "   [--loc \"event location\"]\n"
        "   [--cal \"calendar to be used\"]\n"
        "   [--invite : email(s) to be invited, separated by space]\n"
        "   [--alarm_type : alarm to be set on event: \"DISPLAY\" or \"EMAIL\". Default: none]\n"
        "   [--alarm_format : \"h\" = hours, \"d\" = days]\n"
        "   [--alarm_time : time before the event to set an alarm for]\n"
    )
    return err


# load user settings from file
def load_user_settings(user_conf_json) -> dict:
    logger.debug("Loading user json settings from file: " + str(user_conf_json))
    with open(user_conf_json, 'r') as f:
        return json.load(f)


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
        datetime.strptime(time, "%H:%M:%S")
    except ValueError as err:
        return False, err
    return True, ""


# create ICS file with provided event details
def create_ics(user_settings, event_details) -> None:

    # init calendar
    mycal = Calendar()
    # set properties to be compliant
    mycal.add("prodid", f"-//CalDav ICS Client//{version_num}//{user_settings['domain']}//github.com/ynad/caldav-py-handler//")
    mycal.add("version", "2.0")
    #mycal.add('method', "REQUEST")

    # add calendar subcomponents
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
    myevent.add('created', create_time)
    
    # uid - unique event ID
    myevent['uid'] = event_details['uid']
    myevent.add('priority', 5)

    # add organizer
    organizer = vCalAddress(f"MAILTO:{user_settings['organizer_email']}")
    organizer.params['name'] = vText(user_settings['organizer_name'])
    organizer.params['role'] = vText(user_settings['organizer_role'])
    myevent['organizer'] = organizer

    # add invites if present
    if 'invite' in event_details:
        invitees = event_details['invite'].split()
        for i in invitees:
            attendee = vCalAddress(f"MAILTO:{i}")
            attendee.params['name'] = vText(i)
            attendee.params['role'] = vText('REQ-PARTICIPANT')
            myevent.add('attendee', attendee, encode=0)

    # add an alarm for the event
    if 'alarm_type' in event_details:
        myalarm = Alarm()
        myalarm.add("action", event_details['alarm_type'])
        myalarm.add('summary', event_details['name'])
        myalarm.add('description', event_details['description'])

        # if invitees are present, add email notification
        if 'invite' in event_details:
            invitees = event_details['invite'].split()
            for i in invitees:
                attendee = vCalAddress(f"MAILTO:{i}")
                #attendee.params['name'] = vText(i)
                #attendee.params['role'] = vText('REQ-PARTICIPANT')
                myalarm.add('attendee', attendee, encode=0)
        #myalarm.add("trigger", timedelta(days=-reminder_days))
        # The only way to convince Outlook to do it correctly
        myalarm.add("TRIGGER;RELATED=START", f"-PT{event_details['alarm_time']}{event_details['alarm_format']}")
        myevent.add_component(myalarm)

    # add event to the calendar
    mycal.add_component(myevent)

    # write event to ICS file
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
    with open(ics_file, 'rb') as f:
        data = f.read()

    # if calendar is not set go default
    if calendar == None:
        calendar = user_settings['calendar_default']

    # make PUT request
    try:
        res = requests.put(url=f"{user_settings['server']}/{user_settings['username']}/{calendar}/{event_id}",
                            data=data,
                            headers={'Content-Type': 'text/calendar', 'User-Agent': user_agent},
                            auth = HTTPBasicAuth(user_settings['username'], user_settings['password']))

        if (res.status_code == 201):
            msg = f"Created ({res.status_code})"
            print(msg)
            logger.debug(msg)

        elif (res.status_code == 202):
            msg = f"Accepted ({res.status_code})"
            print(msg)
            logger.debug(msg)

        elif (res.status_code == 204):
            msg = f"No Content ({res.status_code})"
            print(msg)
            logger.debug(msg)

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
            os.remove(ics_file)
        #pass



@click.command()
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


## Main
def main(name, descr, start_day, start_hr, end_day, end_hr, loc, cal, invite, alarm_type, alarm_format, alarm_time):

    # load user settings from json file
    user_settings = load_user_settings(user_conf_json)

    # check command line arguments
    # start and end day are mandatory
    if not start_day or not end_day:
        err = show_syntax(user_settings)
        #logger.error(err)
        print(err)
        input("Press enter to exit.")
        return

    # check start date format
    date_ok, date_err = check_date(start_day)
    if not date_ok:
        err = show_syntax(user_settings)
        logger.error(date_err)
        print(f"{err}\n{date_err}\n")
        input("Press enter to exit.")
        return

    # check end date format
    date_ok, date_err = check_date(end_day)
    if not date_ok:
        err = show_syntax(user_settings)
        logger.error(date_err)
        print(f"{err}\n{date_err}\n")
        input("Press enter to exit.")
        return

    # check time format
    if start_hr and end_hr:
        time_ok, time_err = check_time(start_hr)
        if not time_ok:
            err = show_syntax(user_settings)
            logger.error(time_err)
            print(f"{err}\n{time_err}\n")
            input("Press enter to exit.")
            return

        time_ok, time_err = check_time(end_hr)
        if not time_ok:
            err = show_syntax(user_settings)
            logger.error(time_err)
            print(f"{err}\n{time_err}\n")
            input("Press enter to exit.")
            return

    # build event details
    event_details = {
        'name' : name,
        'description' : descr,
        'calendar' : cal if cal else None,
        'location' : loc if loc else user_settings['location_default'],
        'uid' : (f"{str(datetime.now().timestamp())}_{name}@{user_settings['domain']}").replace(" ", "-")
    }

    # event with fixed hours
    if start_hr and end_hr:
        event_details.update( { 'start' : datetime.strptime(f"{start_day} {start_hr}", "%d/%m/%Y %H:%M:%S") } )
        event_details.update( { 'end' : datetime.strptime(f"{end_day} {end_hr}", "%d/%m/%Y %H:%M:%S") } )
    # full day event
    else:
        event_details.update( { 'start' : datetime.strptime(f"{start_day}", "%d/%m/%Y").date() } )
        event_details.update( { 'end' : datetime.strptime(f"{end_day}", "%d/%m/%Y").date() + timedelta(days=1) } )
    
    # add invitees, can be 1 or more separated by a space
    if invite:
        event_details.update( { 'invite' : invite } )

    # set alarm - all 3 parameters must be given otherwise none is set
    if alarm_type and alarm_type and alarm_time:
        alarm_type = alarm_type.upper()
        alarm_format = alarm_format.upper()
        if ( alarm_type == 'DISPLAY' or alarm_type == 'EMAIL') and ( alarm_format == 'H' or alarm_type == 'D' ):
            event_details.update( { 'alarm_type' : alarm_type.upper() } )
            event_details.update( { 'alarm_format' : alarm_format.upper() } )
            event_details.update( { 'alarm_time' : alarm_time } )

    # wait for user confirmation, if enabled. To skip change 'prompt_wait' to False
    if prompt_wait:
        print(f"\nCaldav ICS CLIent - v{version_num} - {user_settings['domain']}\n"
              "==============================================\n\n"
              f"The following event will be added:\n\n")
        print(f"EVENT NAME:\t{name}\n"
              f"DESCRIPTION:\t{descr}\n"
              f"START DATE:\t{datetime.strftime(event_details['start'], '%d/%m/%Y %H:%M:%S')}\n"
              f"END DATE:\t{datetime.strftime(event_details['end'], '%d/%m/%Y %H:%M:%S')}\n"
              f"LOCATION\t{event_details['location']}\n"
              f"CALENDAR:\t{event_details['calendar']}")
        if invite:
            print(f"INVITEE:\t{invite}")
        if 'alarm_type' in event_details:
            print(f"ALARM:\t\t{event_details['alarm_type']}, {event_details['alarm_time']}{event_details['alarm_format']} before\n")

        input("Press enter to confirm.")

    # compile ICS file
    create_ics(user_settings, event_details)

    # upload it to caldav server
    webdav_put_ics(user_settings, event_details['calendar'], event_details['uid'])

    if prompt_wait:
        input("Press enter to exit.")


if __name__ == '__main__':
    main()
