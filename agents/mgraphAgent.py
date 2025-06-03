#!/usr/bin/python

# Copyright 2025 Daniele Vercelli - ynad <info@danielevercelli.it>
# https://github.com/ynad/calendar-pyhandler
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, version 3 only.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>

"""
# mgraphAgent.py
2025-05-28

Class to handle events creation via Microsoft Graph REST API.
Generates an OAuth access token with interactive user login.
Token cache file and User-Agent may be overridden via properties.
Current capabilities: 
  * Events creation: user calendars, shared calendars, group calendars

'user_settings' dict format:
       {
           "mode" : "microsoft_graph",
           "domain" : "cloud.domain.com",
           "azure_client_id" : "client-id",
           "azure_tenant_id" : "tenant-id",
           "username": "jane.doe",
           "calendar" : "personal",
           "organizer_name" : "Jane Doe",
           "organizer_role" : "IT",
           "organizer_email" : "info@example.com",
           "location" : "Main Office",
           "report" : "path/to/reports-folder"
       }

See README.me for full details.
"""

###################################################################################################
# APP SETTINGS - DO NOT EDIT
VERSION_NUM = "0.7.1"
DEV_EMAIL = "info@danielevercelli.it"
PROD_NAME = "mgraph-pyAgent"
PROD_URL = "github.com/ynad/calendar-pyhandler"
###################################################################################################



import os
import json
import logging
import requests
import msal
from datetime import datetime, timezone



# logger
logger = logging.getLogger(__name__)



class MGraphAgent():

    def __init__(self,
                user_settings: dict,
                user_agent: str = None,
                cache_file: str = "token_cache.json",
        ):
        logger.info("init MGraphAgent")

        # check user settings
        assert 'azure_client_id' in user_settings and 'azure_tenant_id' in user_settings, 'missing Azure user settings'

        self.__user_settings = user_settings
        self.user_agent = user_agent if user_agent else f"{PROD_NAME}/{VERSION_NUM}"

        # Token cache file
        self.cache_file = cache_file

        # Microsoft Graph API Base URL
        self.graph_url = "https://graph.microsoft.com/v1.0"

        self.ms_authority = f"https://login.microsoftonline.com/{self.__user_settings['azure_tenant_id']}"

        # graph API scopes
        self.ms_scopes = [
            "https://graph.microsoft.com/Calendars.ReadWrite", 
            "https://graph.microsoft.com/Calendars.ReadWrite.Shared", 
            "https://graph.microsoft.com/Group.ReadWrite.All", 
            "https://graph.microsoft.com/offline_access"
        ]

        self.access_token = self.__get_access_token()


    def __get_access_token(self) -> str:
        """ Retrieves access token from cache or authenticates user if needed """
        logger.info("get_access_token")
        
        # Load token cache if exists
        cache = msal.SerializableTokenCache()
        if os.path.exists(self.cache_file):
            with open(self.cache_file, "r") as fp:
                cache.deserialize(fp.read())
            logger.info(f"token read from cache file")

        graph_app = msal.PublicClientApplication(self.__user_settings['azure_client_id'], authority=self.ms_authority, token_cache=cache)

        # Try to get token silently (without asking user)
        accounts = graph_app.get_accounts()
        if accounts:
            logger.info(f"acquire token silent")
            result = graph_app.acquire_token_silent(self.ms_scopes, account=accounts[0])

        else:
            # If no valid token, ask the user to log in interactively
            logger.info(f"interactive user log in")
            result = graph_app.acquire_token_interactive(self.ms_scopes)

        if "access_token" in result:
            logger.info(f"access_token found")
            # Save updated token cache
            with open(self.cache_file, "w") as fp:
                fp.write(cache.serialize())
            
            return result["access_token"]

        else:
            err = "Failed to authenticate: " + json.dumps(result, indent=4)
            logger.error(err)
            raise Exception(err)


    def __request_post(self, url: str, payload: dict) -> tuple[bool, str]:
        logger.info(f"request POST, url endpoint: {url}, payload: {payload}")
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "User-Agent": self.user_agent
        }
        response = requests.post(url, headers=headers, json=payload)
        logger.debug(f"request headers: {response.request.headers}")
        logger.debug(f"response headers: {response.headers}")
    
        if response.status_code == 201:
            msg = f"Event created ({response.status_code})"
            logger.info(msg)
        else:
            msg = f"ERROR: {response.status_code}, {response.reason}: {response.text}"
            logger.error(msg)

        return bool(response.status_code == 201), msg


    def create_event(self, event_data: dict) -> tuple[bool, str]:
        # prepare event json
        event_details = self.__format_event(event_data)

        # endpoint: personal default calendar
        if event_data['calendar'] == 'personal':
            url = f"{self.graph_url}/me/events"
        # group calendar
        elif 'group' in event_data and event_data['group']:
            url = f"{self.graph_url}/groups/{event_data['calendar']}/events"
        # other calendars (personal/shared)
        else:
            url = f"{self.graph_url}/me/calendars/{event_data['calendar']}/events"

        # send to ms graph API
        res, msg = self.__request_post(url, event_details)

        # append result message
        msg = f"{event_data['name']}\n{msg}"
        print(msg)

        return res, msg


    def __format_event(self, event_details: dict) -> dict:
        # base date
        event_data = {
            "subject": event_details['name'],
            "start": {
                "dateTime": datetime.strftime(event_details['start'], '%Y-%m-%dT%H:%M:%S'), 
                "timeZone": "Europe/Berlin"
            },
            "end": {
                "dateTime": datetime.strftime(event_details['end'], '%Y-%m-%dT%H:%M:%S'),
                "timeZone": "Europe/Berlin"
            },
            "body": {
                "content": event_details['description'],
                "contentType": "text"
            },
            "createdDateTime" : datetime.now(timezone.utc).isoformat(),
            "importance" : "normal"
        }
        # bool flag for full day event
        if event_details['fullday']:
            event_data['isAllDay'] = True

        # location
        if 'location' in event_details and event_details['location']:
            event_data['location'] = {
                "displayName": event_details['location']
            }

        # list of attendees
        if 'invite' in event_details:
            invitees = event_details['invite'].split()
            event_data['attendees'] = []
            for i in invitees:
                event_data['attendees'].append(
                    { 
                        "emailAddress" :
                        { 
                            "address" : i
                        },
                        "type" : "required" 
                    }
                )

        # optional organizer info
        if 'organizer_email' in self.__user_settings:
            if 'organizer_name' in self.__user_settings:
                if 'organizer_role' in self.__user_settings:
                    name = f"{self.__user_settings['organizer_name']} ({self.__user_settings['organizer_role']})"
                else:
                    name = f"{self.__user_settings['organizer_name']}"
            else:
                name = ''

            event_data['organizer'] = { 
                        "emailAddress" :
                        {
                            "address" : self.__user_settings['organizer_email'],
                            "name" : name
                        }
                    }

        # add an alarm for the event
        if 'alarm_type' in event_details:
            if event_details['alarm_type'] == 'DISPLAY':

                try:
                    # calc time in minutes
                    if event_details['alarm_format'] == 'D':
                        reminder_mins = int(event_details['alarm_time']) * 1440

                    elif event_details['alarm_format'] == 'H':
                        h, m = event_details['alarm_time'].split(':')
                        reminder_mins = (int(h) * 60) + int(m)
                except ValueError as exc:
                    logger.warning(f"Invalid alarm_time, must be integer HH:MM or D")
                    raise ValueError(f"Invalid alarm_time, must be integer HH:MM or D")

                logger.info(f"alarm reminder set to: {reminder_mins} minutes")
                event_data['reminderMinutesBeforeStart'] = reminder_mins
                event_data['isReminderOn'] = True

            elif event_details['alarm_type'] == 'EMAIL':
                logger.warning("Alarm EMAIL not implemented")
                raise ValueError("Alarm EMAIL not implemented")
            else:
                logger.error("Invalid alarm_type")
                raise ValueError("Invalid alarm_type")


        logger.debug(f"event data formatted: {event_data}")
        return event_data


