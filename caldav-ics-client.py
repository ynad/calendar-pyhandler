#!/usr/bin/python

#
# caldav-ics-client
# receives event details from arguments, parse them, create an ICS event, upload it via webdav PUT request
# must provide JSON formatted: User-Agent, webdav server url, authentication
# based on a NextCloud environment
#
# v0.4 - 2022.11.11 - https://github.com/ynad/caldav-py-handler
# info@danielevercelli.it
#

import sys, os, logging
import json
import click
import requests
from typing import Dict, List, Tuple
from requests.auth import HTTPBasicAuth
from icalendar import Calendar, Event, vCalAddress, vText
from datetime import datetime, timedelta
from pathlib import Path



###################################################################################################
# USER SETTINGS - adjust to your environment
domain="cloud.domain.com"
user_conf_json="user_settings.json"
logging_file="debug.log"
prompt_wait=True

# APP SETTINGS - no need to edit normally
user_agent="caldav-ics-client"
ics_file="tmp-event.ics"
###################################################################################################



# Enable logging
logging.basicConfig(
    filename=logging_file,
    filemode='w',
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)
logger = logging.getLogger(__name__)



# show command syntax
def show_syntax() -> str:
    err = (
        f"\nCaldav ICS CLIent - {domain}\n"
        "=======================================\n\n"
        f"Missing or wrong arguments! Syntax:\n{sys.argv}\n"
        "    --name \"event name\"\n"
        "    --descr \"event description\"\n"
        "    --start_day \"event start day\"\n"
        "    --end_day \"event end day\"\n"
        "   [--start_hr \"event start hour\"]\n"
        "   [--end_hr \"event end hour\"]\n"
        "   [--loc \"event location\"]\n"
        "   [--cal \"calendar to be used\"]\n"
        "   [--invitee \"email to be invited\"]\n"
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
        datetime.strptime(date, "%m/%d/%Y")
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
    cal = Calendar()
    # set properties to be compliant
    cal.add('prodid', f'-//CalDav ICS Client//{domain}//github.com/ynad/caldav-py-handler//')
    cal.add('version', '2.0')

    # add calendar subcomponents
    event = Event()
    #event.add('name', event_details['name'])
    event.add('summary', event_details['name'])
    event.add('description', event_details['description'])
    event.add('dtstart', event_details['start'])
    event.add('dtend', event_details['end'])

    # add organizer
    organizer = vCalAddress(user_settings['organizer_email'])
    organizer.params['name'] = vText(user_settings['organizer_name'])
    organizer.params['role'] = vText(user_settings['organizer_role'])
    event['organizer'] = organizer

    # add location
    event['location'] = vText(event_details['location'])
    # uid - unique event ID
    event['uid'] = event_details['uid']
    event.add('priority', 5)

    # add invites if present
    if 'invite' in event_details:
        # add attendees - to-do list of attendees
        attendee = vCalAddress(event_details['invite'])
        attendee.params['name'] = vText(event_details['invite'])
        attendee.params['role'] = vText('REQ-PARTICIPANT')
        event.add('attendee', attendee, encode=0)

    # add event to the calendar
    cal.add_component(event)

    # write event to ICS file
    with open(ics_file, 'wb') as f:
        f.write(cal.to_ical())


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
    default="Main Office"
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

## Main
def main(name, descr, start_day, start_hr, end_day, end_hr, loc, cal, invite):

    # check command line arguments
    if not start_day or not end_day:
        err = show_syntax()
        #logger.error(err)
        print(err)
        input("Press any key to exit.")
        return

    # check date format
    date_ok, date_err = check_date(start_day)
    if not date_ok:
        err = show_syntax()
        logger.error(date_err)
        print(f"{err}\n{date_err}\n")
        input("Press any key to exit.")
        return

    date_ok, date_err = check_date(end_day)
    if not date_ok:
        err = show_syntax()
        logger.error(date_err)
        print(f"{err}\n{date_err}\n")
        input("Press any key to exit.")
        return

    # check time format
    if start_hr and end_hr:
        time_ok, time_err = check_time(start_hr)
        if not time_ok:
            err = show_syntax()
            logger.error(time_err)
            print(f"{err}\n{time_err}\n")
            input("Press any key to exit.")
            return

        time_ok, time_err = check_time(end_hr)
        if not time_ok:
            err = show_syntax()
            logger.error(time_err)
            print(f"{err}\n{time_err}\n")
            input("Press any key to exit.")
            return

    # load user settings from json file
    user_settings = load_user_settings(user_conf_json)

    # build event details
    event_details = {
        'name' : name,
        'description' : descr,
        'calendar' : cal if cal else None,
        'location' : loc,
        'uid' : (f"{str(datetime.now().timestamp())}_{name}@{domain}").replace(" ", "-")
    }
    # event with fixed hours
    if start_hr and end_hr:
        event_details.update({'start' : datetime.strptime(f"{start_day} {start_hr}", "%d/%m/%Y %H:%M:%S")})
        event_details.update({'end' : datetime.strptime(f"{end_day} {end_hr}", "%d/%m/%Y %H:%M:%S")})
    # full day event
    else:
        event_details.update({'start' : datetime.strptime(f"{start_day}", "%d/%m/%Y").date()})
        event_details.update({'end' : datetime.strptime(f"{end_day}", "%d/%m/%Y").date() + timedelta(days=1)})
    # add invitees - to-do list of invitations
    if invite:
        event_details.update({'invite' : invite})

    # wait for user confirmation, if enabled. To skip change 'prompt' to False
    if prompt_wait:
        print(f"\nCaldav ICS CLIent - {domain}\n"
              "=======================================\n\n"
              f"The following event details were provided:\n\n")
        print(f"EVENT NAME:\t{name}\n"
              f"DESCRIPTION:\t{descr}\n"
              f"START DATE:\t{datetime.strftime(event_details['start'], '%d/%m/%Y %H:%M:%S')}\n"
              f"END DATE:\t{datetime.strftime(event_details['end'], '%d/%m/%Y %H:%M:%S')}\n"
              f"LOCATION\t{loc}\n"
              f"CALENDAR:\t{event_details['calendar']}\n")
        input("Press any key to confirm.")

    # compile ICS file
    create_ics(user_settings, event_details)

    # upload it to caldav server
    webdav_put_ics(user_settings, event_details['calendar'], event_details['uid'])

    input("Press any key to exit.")


if __name__ == '__main__':
    main()
