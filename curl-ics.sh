#!/bin/bash

#
# sample curl command - send webdav PUT request to upload binary ICS file to given calendar
# must provide authentication, User-Agent and Content-Type
# based on a NextCloud test environment
#
# v0.1 - 2022.10.24 - https://github.com/ynad/caldav-py-handler
# info@danielevercelli.it
#

curl --user 'username:password' -X PUT \
https://cloud.domain/remote.php/dav/calendars/username/personal/curl-event.ics \
-H "User-Agent: ics-test" \
-H "Content-Type: text/calendar ; charset=utf-8" --data-binary @curl-event.ics
