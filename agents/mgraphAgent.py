#!/usr/bin/python

#
## mgraphAgent.py
# Class to handle events creation via Microsoft Graph REST API.
# Generates an OAuth access token with interactive user login.
# Token cache file and User-Agent may be overridden via properties.
# Current capabilities: 
#   * Events creation: user calendars, shared calendars, group calendars
#
# 'user_settings' dict format:
#        {
#            "mode" : "microsoft_graph",
#            "domain" : "cloud.domain.com",
#            "azure_client_id" : "client-id",
#            "azure_tenant_id" : "tenant-id",
#            "username": "jane.doe",
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
PROD_NAME = "mgraph-pyAgent"
PROD_URL = "github.com/ynad/calendar-pyhandler"
###################################################################################################



import os
import json
import logging
import requests
import msal
from datetime import datetime, timedelta



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

        # Save updated token cache
        with open(self.cache_file, "w") as fp:
            fp.write(cache.serialize())

        if "access_token" in result:
            logger.info(f"access_token found")
            return result["access_token"]

        else:
            err = "Failed to authenticate: " + json.dumps(result, indent=4)
            logger.error(err)
            raise Exception(err)


    def __request_post(self, url: str, payload: dict) -> tuple[bool, str]:
        logger.info(f"request POST, url endpoint: {url}, event_data: {event_data}")
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
        res, msg = self.__post_event(url, event_details)

        # append result message
        msg = f"{event_data['name']}\n{msg}"
        print(msg)

        return res, msg


    def __format_event(self, event_data: dict) -> dict:
        event_details = {
            "subject": event_data['name'],
            "start": {
                "dateTime": datetime.strftime(event_data['start'], '%Y-%m-%dT%H:%M:%S'), 
                "timeZone": "Europe/Berlin"
            },
            "end": {
                "dateTime": datetime.strftime(event_data['end'], '%Y-%m-%dT%H:%M:%S'),
                "timeZone": "Europe/Berlin"
            },
            "body": {
                "content": event_data['description'],
                "contentType": "text"
            }
        }
        if 'location' in event_data and event_data['location']:
            event_details['location'] = {
                "displayName": event_data['location']
            }

        if 'invite' in event_data:
            invitees = event_data['invite'].split()
            event_details['attendees'] = []
            for i in invitees:
                event_details['attendees'].append(
                    { 
                        "emailAddress":
                        { 
                            "address": i
                        },
                        "type": "required" 
                    }
                )

        logger.debug(f"event data formatted: {event_details}")

        return event_details


