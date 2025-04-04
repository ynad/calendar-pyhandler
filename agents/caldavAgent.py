#!/usr/bin/python

#
## caldavAgent.py
# Class to handle CalDav (WebDAV) requests.
# ICS file and User-Agent may be overridden via properties.
# Current capabilities: 
#   * Events creation: Writes an ICS file with event details and send PUT request
#
# 'user_settings' dict format:
#        {
#            "mode" : "caldav",
#            "domain" : "cloud.domain.com",
#            "server" : "https://cloud.domain.com/remote.php/dav/calendars",
#            "username": "jane.doe",
#            "password": "secret-app-password",
#            "calendar" : "personal",
#            "organizer_name" : "Jane Doe",
#            "organizer_role" : "IT",
#            "organizer_email" : "info@example.com",
#            "location" : "Main Office",
#            "report" : "path/to/reports-folder"
#        }
#
# See README.me for full details.
#
## License
# Released under GPL-3.0 license.
#
# 2025.04.04
# https://github.com/ynad/calendar-pyhandler
# info@danielevercelli.it
#

###################################################################################################
# APP SETTINGS - DO NOT EDIT
VERSION_NUM = "0.6.0"
DEV_EMAIL = "info@danielevercelli.it"
PROD_NAME = "CalDav-pyAgent"
PROD_URL = "github.com/ynad/calendar-pyhandler"
###################################################################################################



import os
import logging
import requests
from datetime import datetime, timedelta
from icalendar import Calendar, Event, Alarm, vCalAddress, vText
from requests.auth import HTTPBasicAuth



# logger
logger = logging.getLogger(__name__)



class CaldavAgent():

    def __init__(self,
                user_settings: dict,
                ics_file: str = 'tmp_caldav-event.ics',
                user_agent: str = None,
        ):
        logger.info("init CaldavAgent")

        self.__user_settings = user_settings
        self.ics_file = ics_file
        self.user_agent = user_agent if user_agent else f"{PROD_NAME}/{VERSION_NUM}"


    def create_event(self, event_data: dict) -> tuple[bool, str]:
        # compile ICS file
        self.__create_ics(event_data)

        # upload it to caldav server
        res, msg = self.__webdav_put_ics(event_data['calendar'], event_data['uid'])

        # append result message
        msg = f"{event_data['name']}\n{msg}"

        return res, msg


    # create ICS file with provided event details
    def __create_ics(self, event_details: dict) -> None:
        # init calendar
        logger.info(f"ICS: create calendar")
        mycal = Calendar()

        # set properties to be compliant
        mycal.add("prodid", f"-//{PROD_NAME}//{VERSION_NUM}//{self.__user_settings['domain']}//{PROD_URL}//")
        mycal.add("version", "2.0")
        #mycal.add('method', "REQUEST")

        # add calendar subcomponents
        logger.info(f"ICS: create event")
        myevent = Event()
        #myevent.add('name', event_details['name'])
        myevent.add('summary', event_details['name'])
        myevent.add('description', event_details['description'])
        myevent.add('dtstart', event_details['start'])
        myevent.add('dtend', event_details['end'])
        myevent.add('status', "confirmed")

        # add location
        if 'location' in event_details and event_details['location']:
            myevent['location'] = vText(event_details['location'])

        # creation time
        create_time = datetime.strptime(datetime.now().strftime("%d/%m/%Y %H:%M:%S"), "%d/%m/%Y %H:%M:%S")
        logger.info(f"ICS: created time: {create_time}")
        myevent.add('created', create_time)
        
        # uid - unique event ID
        myevent['uid'] = event_details['uid']
        myevent.add('priority', 5)

        # add organizer
        logger.info(f"ICS: organizer: {self.__user_settings['organizer_email']}")
        organizer = vCalAddress(f"MAILTO:{self.__user_settings['organizer_email']}")
        organizer.params['CN'] = vText(self.__user_settings['organizer_name'])
        organizer.params['role'] = vText(self.__user_settings['organizer_role'])
        myevent['organizer'] = organizer

        # add invites if present
        if 'invite' in event_details:
            invitees = event_details['invite'].split()
            for i in invitees:
                logger.info(f"ICS: adding invite for: {i}")
                attendee = vCalAddress(f"MAILTO:{i}")
                attendee.params['CN'] = vText(i)
                attendee.params['role'] = vText('REQ-PARTICIPANT')
                attendee.params['PARTSTAT'] = vText('NEEDS-ACTION')
                attendee.params['RSVP'] = vText('TRUE')
                myevent.add('attendee', attendee, encode=0)

        # add an alarm for the event
        if 'alarm_type' in event_details:
            logger.info(f"ICS: adding alarm")
            myalarm = Alarm()
            myalarm.add("action", event_details['alarm_type'])
            myalarm.add('summary', event_details['name'])
            myalarm.add('description', event_details['description'])

            # if invitees are present, add email notification
            if 'invite' in event_details:
                invitees = event_details['invite'].split()
                for i in invitees:
                    logger.info(f"ICS: alarm invite for: {i}")
                    attendee = vCalAddress(f"MAILTO:{i}")
                    #attendee.params['name'] = vText(i)
                    #attendee.params['role'] = vText('REQ-PARTICIPANT')
                    myalarm.add('attendee', attendee, encode=0)

            # set trigger time
            #myalarm.add("trigger", timedelta(days=-reminder_days))
            if event_details['fullday'] is True:
                logger.info(f"ICS: fullday event")
                myalarm.add("TRIGGER;RELATED=START", f"-P{event_details['alarm_time']}{event_details['alarm_format']}")
            else:
                logger.info(f"ICS: fixed hours event")
                myalarm.add("TRIGGER;RELATED=START", f"-PT{event_details['alarm_time']}{event_details['alarm_format']}")
            myevent.add_component(myalarm)

        # add event to the calendar
        mycal.add_component(myevent)

        # write event to ICS file
        logger.info(f"ICS: write to file {self.ics_file}")
        with open(self.ics_file, 'wb') as f:
            f.write(mycal.to_ical())


    # make PUT request to upload ICS event file to given calendar
    def __webdav_put_ics(self, calendar: str, event_id: str) -> tuple[bool, str]:
        # check ICS existance
        if not os.path.exists(self.ics_file):
            err = f"ERROR: missing ICS file {self.ics_file}, can't continue"
            print(err)
            logger.error(err)
            return False

        # read ICS ifle
        logger.info(f"webdav: read ics from file {self.ics_file}")
        with open(self.ics_file, 'rb') as f:
            data = f.read()

        # if calendar is not set go default
        if calendar == None:
            calendar = 'personal'
        logger.info(f"webdav: calendar set to: {calendar}")

        # make PUT request
        try:
            headers = {
                'Content-Type': 'text/calendar', 
                'User-Agent': self.user_agent
            }
            logger.info(f"webdav: put request headers: {headers}")

            res = requests.put(url=f"{self.__user_settings['server']}/{self.__user_settings['username']}/{calendar}/{event_id}",
                                data=data,
                                headers=headers,
                                auth = HTTPBasicAuth(self.__user_settings['username'], self.__user_settings['password']))

            if (res.status_code == 201):
                msg = f"Event created ({res.status_code})"
                print(msg)
                logger.info(msg)

            elif (res.status_code == 202):
                msg = f"Event accepted ({res.status_code})"
                print(msg)
                logger.info(msg)

            elif (res.status_code == 204):
                msg = f"No Content ({res.status_code})"
                print(msg)
                logger.info(msg)

            else:
                msg = (
                    f"ERROR: {res.status_code}, {res.reason}: {res.text}"
                )
                raise Exception(msg)

            put_ack = True

        except Exception as exc:
            print(exc)
            logger.error(exc)
            put_ack = False

        finally:
            # rm tmp ics files
            if os.path.exists(self.ics_file):
                logger.info(f"webdav: remove ics file {self.ics_file}")
                os.remove(self.ics_file)
            #pass

        return put_ack, msg


