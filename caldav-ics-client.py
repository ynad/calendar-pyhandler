#!/usr/bin/python

#
# caldav-ics-client
# receives event details from arguments, parse them, create an ICS event, upload it via webdav PUT request
# must provide User-Agent, webdav server url, authentication
# based on a NextCloud test environment
#
# v0.2 - 2022.10.25 - https://github.com/ynad/caldav-py-handler
# info@danielevercelli.it
#

import sys, os
import logging
import json
import pytz
import requests
from requests.auth import HTTPBasicAuth

from icalendar import Calendar, Event, vCalAddress, vText
from datetime import datetime
from pathlib import Path


# Enable logging
logging.basicConfig(
    filename="./debug.log",
    filemode='w',
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)
logger = logging.getLogger(__name__)


# global settings
user_agent='caldav-ics-client'
ics_file='test_caldav.ics'
event_ics='test-caldav-event.ics'

# user settings
user_conf_json='user_settings.json'


# load user settings from file
def load_user_settings(user_conf_json):
    logger.debug("Loading user json settings from file: " + str(user_conf_json))
    with open(user_conf_json, 'r') as f:
        return json.load(f)


# create ICS file with provided event details
def create_ics(event_details):

    # init calendar
    cal = Calendar()
    # set properties to be compliant
    cal.add('prodid', '-//CalDav ICS Client//cloud.domain.com//github.com/ynad/caldav-py-handler//')
    cal.add('version', '2.0')

    # add calendar subcomponents
    event = Event()
    #event.add('name', event_details['name'])
    event.add('summary', event_details['name'])
    event.add('description', event_details['description'])
    event.add('dtstart', datetime(2022, 10, 25, 8, 0, 0, tzinfo=pytz.utc))
    event.add('dtend', datetime(2022, 10, 25, 10, 0, 0, tzinfo=pytz.utc))

    # add organizer
    organizer = vCalAddress('MAILTO:info@danielevercelli.it')

    # add event parameters
    organizer.params['name'] = vText('Foo Bar')
    organizer.params['role'] = vText('CEO')
    event['organizer'] = organizer
    event['location'] = vText('Roma, IT')

    event['uid'] = '2022125T111010/272356262376@example.com'
    event.add('priority', 5)

    # add attendees - to-do list of attendees
    attendee = vCalAddress('MAILTO:info@danielevercelli.it')
    attendee.params['name'] = vText('Micky Mouse')
    attendee.params['role'] = vText('REQ-PARTICIPANT')
    event.add('attendee', attendee, encode=0)

    # add event to the calendar
    cal.add_component(event)

    # write event to ICS file
    with open(ics_file, 'wb') as f:
        f.write(cal.to_ical())


# make PUT request to upload ICS event file to given calendar
def webdav_put_ics(user_settings, calendar, event_ics):

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
        res = requests.put(url=f"{user_settings['server']}/{user_settings['username']}/{calendar}/{event_ics}",
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


## Main
def main():

    # check command line arguments
    if len(sys.argv) < 5:
        err = (
            f"Missing arguments!\n"
            + f"Syntax:\n{sys.argv} \"event name\" \"event description\" \"event start time\" \"event end time\" \"calendar name\""
        )
        print(err)
        logger.error(err)
        return

    # load user settings from json file
    user_settings = load_user_settings(user_conf_json)

    # acquire event settings from args
    if len(sys.argv) == 6:
       calendar = sys.argv[5]
    else:
        calendar = None
    event_details = {
        'name' : sys.argv[1],
        'description' : sys.argv[2],
        'start' : sys.argv[3],
        'end' : sys.argv[4],
        'calendar' : calendar
    }

    create_ics(event_details)

    webdav_put_ics(user_settings, event_details['calendar'], event_ics)


if __name__ == '__main__':
    main()
