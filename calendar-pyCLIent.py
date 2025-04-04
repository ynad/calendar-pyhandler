#!/usr/bin/python

#
## calendar-pyCLIent.py
# Receives event details from arguments, parse them, create an ICS event, upload it via webdav PUT request.
# Must provide JSON formatted file with: User-Agent, webdav server url, authentication, etc.
# Based on a NextCloud environment.
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
PROD_NAME = "calendar-pyCLIent"
PROD_URL = "github.com/ynad/calendar-pyhandler"
update_version_url = "https://raw.githubusercontent.com/ynad/calendar-pyhandler/main/VERSION"
update_url = "https://raw.githubusercontent.com/ynad/calendar-pyhandler/main/calendar-pyCLIent-client.py"
requirements_url = "https://raw.githubusercontent.com/ynad/calendar-pyhandler/main/requirements.txt"
logging_file = "debug.log"
requirements_file = "requirements.txt"
pip_json = "pip.json"
ics_file = "tmp_calendar-pyhandler-event.ics"
###################################################################################################



import sys
import os 
import shutil
import logging
import json
import click
import requests
import urllib.request
import random
import signal
from datetime import datetime, timedelta
from pathlib import Path

# GUI libs
import tkinter as tk
from tkinter import scrolledtext, messagebox, Toplevel

# internal libs
from agents.caldavAgent import CaldavAgent
from agents.mgraphAgent import MGraphAgent



# Enable logging
logging_file = f"{os.path.dirname(__file__)}/{logging_file}"
logging.basicConfig(
    filename=logging_file,
    filemode='w',
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.DEBUG
)
logger = logging.getLogger(__name__)

# set filenames on local path
# requirements
requirements_file=f"{os.path.dirname(__file__)}/{requirements_file}"
# pip list
pip_json=f"{os.path.dirname(__file__)}/{pip_json}"
# ics tmp file
ics_file=f"{os.path.dirname(__file__)}/{ics_file}"



def message_box(msg_text: str, msg_type: str = 'info') -> None:
    window = tk.Tk()
    window.wm_withdraw()

    if msg_type == 'error':
        messagebox.showerror(title=f"Error - {string_header(short=True)}", message=msg_text)
    elif msg_type == 'warning':
        messagebox.showwarning(title=f"Warning - {string_header(short=True)}", message=msg_text)
    else:
        messagebox.showinfo(title=f"Information - {string_header(short=True)}", message=msg_text)

    window.destroy()
    return None


def confirm_box(output_tk: str, events_list: list) -> None:
    # Create main window
    root = tk.Tk()
    root.title(string_header(terminal=False))
    root.geometry("900x600")

    # Label for instruction or title
    label = tk.Label(root, text=f"\nThe following event(s) will be created:\n")
    label.pack(pady=10)

    # ScrolledText widget for displaying the text
    text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=20)
    text.pack(pady=10)

    # Sample text for demonstration (between 15 and 50 lines)
    text.insert(tk.END, output_tk)
    text.configure(state='disabled')  # Make the text read-only

    # Button frame to organize Confirm and Cancel buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=20)

    # Confirm and Cancel buttons
    confirm_button = tk.Button(button_frame, text="OK", command=lambda: [create_events(events_list), root.destroy(), root.quit()])
    cancel_button = tk.Button(button_frame, text="Cancel", command=lambda: [root.destroy(), root.quit()])

    confirm_button.pack(side=tk.LEFT, padx=10)
    cancel_button.pack(side=tk.RIGHT, padx=10)

    # Run the application
    root.mainloop()


# determine user backend mode and create events accordingly
def create_events(events_list: list) -> bool:
    logger.info(f"create_events, mode: {user_settings['mode']}")

    # CalDav - WebDav
    if user_settings['mode'] == 'caldav':
        agent = CaldavAgent(user_settings, ics_file=ics_file)

    # Microsoft Graph REST API
    elif user_settings['mode'] == 'microsoft_graph':
        agent = MGraphAgent(user_settings)

    else:
        logger.error(f"invalid client mode: {user_settings['mode']}, cannot continue")
        return False

    for event_n in events_list:
        res, msg = agent.create_event(event_n)
        if res:
            message_box(msg, msg_type='info')
        else:
            message_box(msg, msg_type='error')


# show command syntax
def string_header(terminal: bool = False, short: bool = False) -> str:
    header = f"Calendar pyCLIent" 
    if not short:
        header +=  f" - v{VERSION_NUM} - {user_settings['domain']}"
    if terminal:
        header += "\n==============================================\n"
    return header


# show command syntax
def show_syntax() -> str:
    return(f"Parameters given:\n\n{sys.argv}\n\n\nSyntax:\n\n"
            "Event and calendar settings:\n"
            "    --name \"event name\"\n"
            "    --descr \"event description\"\n"
            "    --start_day dd/mm/YYYY [dd/mm/YYYY [...]]\n"
            "    --end_day dd/mm/YYYY [dd/mm/YYYY [...]]\n"
            "   [--start_hr HH:MM [HH:MM [...]]]\n"
            "   [--end_hr HH:MM [HH:MM [...]]]\n"
            "   [--loc \"event location\"]\n"
            "   [--cal \"calendar to be used\". Default: \"personal\"]\n"
            "   [--invite : email(s) to be invited, separated by space]\n"
            "\nAlarm settings, all 3 parameters must be set or none is considered:\n"
            "   [--alarm_type : \"DISPLAY\" or \"EMAIL\". Alarm to be set on event. Default: none]\n"
            "   [--alarm_format : \"h\" = hours, \"d\" = days]\n"
            "   [--alarm_time : time before the event to set an alarm for, in given format]\n"
            "\nApp behavior settings:\n"
            "   [--config \"path\\to\\config-file.json\". Default: \"user_settings.json\"]\n"
            "   [--noprompt : skip user confirmation]\n"
            "   [--noreport : skip report log copy for developer]\n"
            "   [--noupdate : skip software updates auto-check]\n"
    )


# load user settings from file
def load_user_settings(user_config: str) -> dict:
    logger.info(f"Calendar pyCLIent - v{VERSION_NUM}")
    if os.path.exists(user_config):
        logger.info("Loading user settings JSON from file: " + str(user_config))
        with open(user_config, 'r') as f:
            user_settings = json.load(f)
        logger.info(f"Running instance for: {user_settings['domain']}, user: {user_settings['username']}, calendar: {user_settings['calendar'] if 'calendar' in user_settings else 'None'}")
        return user_settings
    else:
        logger.error(f"User settings JSON missing: {user_config}")
        return None


# send report of usage to developer
def send_report() -> None:
    # copy log to report dir, if path is provided in user_settings
    if 'report' in user_settings:
        try:
            log_report = f"{user_settings['report']}//calendar-pyCLIent_debug_{str(datetime.now().timestamp())}.log"
            shutil.copyfile(logging_file, log_report)
            logger.info(f"Report sent, log copied to {log_report}")

        except Exception as e:
            logger.error(f"Cannot copy debug log file from: {logging_file} to: {log_report}, error: {e}")
    else:
        logger.warning(f"No report path in user_settings, cannot send report")


def check_dependencies() -> bool:
    # run dep check only if an update_version is found
    if update_version:

        # get updated requirements, else use local if exist
        get_requirements()
        if os.path.exists(requirements_file):
            # get os pip package list as json
            logger.info(f"Save local pip list on file: {pip_json}")
            try:
                os.system(f"pip list --disable-pip-version-check --format json > {pip_json}")
                with open(pip_json, 'r') as fp:
                    pip_list = json.load(fp)
                os.remove(pip_json)
            except Exception as e:
                logger.error(f"Error reading pip list JSON from file {pip_json}: {e}")
                print("(!) Error reading software state list.\n")
                return False

            logger.info(f"Read requirements from file: {requirements_file}")
            try:
                with open(requirements_file, 'r') as fp:
                    requirements = fp.readlines()
            except Exception as e:
                logger.error(f"Error reading requirements from file: {requirements_file}")
                print("(!) Error reading software requirements.\n")
                return False

            # iterate over package list and search for my requirements    
            for pack in pip_list:
                for req in requirements:
                    if pack['name'] == req.split('\n')[0]:
                        requirements.remove(req)

            # if there are missing requirements install them
            if len(requirements) > 0:
                return install_requirements(requirements)
            else:
                logger.info(f"All requirements satisfied")
                return True

        else:
            logger.warning(f"Requirements not found on path: {requirements_file}")
            print("(!) Error checking software requirements.\n")
            return False

    else:
        logger.info("check_dependencies skipped")


def get_requirements() -> bool:
    try:
        logger.info(f"Get updated requirements from url: {requirements_url}, to file: {requirements_file}")
        urllib.request.urlretrieve(requirements_url, requirements_file)
        return True

    except Exception as e:
        logger.warning(f"Error downloading python requirements from url {requirements_url}: {e}")
        print("(!) Error occurred while downloading software requirements.\n")
        return False


def install_requirements(requirements: list) -> bool:
    print(f"New software dependencies are available:")
    for req in requirements:
        req = req.split('\n')[0]
        print(f"{req} ")
    print( "\nATTENTION: new python packages will be installed. If skipped, the software might not work after an update.\n"
           "Do you want to install them now? (Y/N)"
           )
    run_update = input()

    if run_update.upper() == 'Y':
        logger.info(f"Installing missing requirements from file: {requirements_file}")
        print(f"Downloading new python pip packages...\n")
        return os.system(f"pip install --disable-pip-version-check -r {requirements_file}")
    else:
        logger.info("Requirements update skipped")
        print("Requirements update skipped.\n")
        return False


# check software updates
def check_updates() -> tuple[str, str]:
    # get latest version number
    response = requests.get(update_version_url)
    if (response.status_code == 200):
        update_info = response.text.split('\n')

        # newer version available on repo
        if (update_info[0]) > (VERSION_NUM):
        #if version.parse(update_info[0]) > version.parse(VERSION_NUM):
            logger.info(f"Current version: {VERSION_NUM}, found update: {update_info}")
            return update_info[0], update_info[1]

        else:
            logger.info(f"No updates available. Current version: {VERSION_NUM}, online version: {update_info}")
            return None, None

    else:
        logger.error(f"{e} {response.status_code}, {response.text}")
        e = "(!) Error occurred while checking available updates."
        print(e)
        message_box(e, msg_type='warning')
        return None, None


# self update app
def self_update() -> None:
    # if a newer version is available on repo
    if update_version:
        print(f"A new version is available: {update_version}, {update_date}\n"
               "After the update you may have to re-launch the program.\n"
               "Do you want to update now? (Y/N)")
        run_update = input()

        if run_update.upper() == 'Y':
            logger.info("Downloading new version and restarting")
            print(f"Downloading new version and restarting...")
            urllib.request.urlretrieve(update_url, __file__)
            os.execl(sys.executable, 'python', __file__, *sys.argv[1:])
        else:
            logger.info("Update skipped")
            print("Update skipped.")

    else:
        logger.info(f"self_update nothing to do")


# check string date format
def check_date(date: str) -> tuple[bool, str]:
    try:
        datetime.strptime(date, "%d/%m/%Y")
    except ValueError as err:
        return False, err
    return True, ""


# check string time format
def check_time(time: str) -> tuple[bool, str]:
    try:
        datetime.strptime(time, "%H:%M")
    except ValueError as err:
        return False, err
    return True, ""


# check if first date is after the second one
def is_after_date(date_i: str, date_j: str) -> bool:
    return datetime.strptime(date_i, "%d/%m/%Y") > datetime.strptime(date_j, "%d/%m/%Y")


# check if first hour is after the second one
def is_after_hour(date_i: str, date_j: str) -> bool:
    return datetime.strptime(date_i, "%H:%M") > datetime.strptime(date_j, "%H:%M")


# check arguments and return error strings
def args_check(start_day: str, end_day: str, start_hr: str, end_hr: str) -> tuple[bool, str]:
    logger.info(f"Running args check")

    # start and end day are mandatory
    if not start_day or not end_day:
        return False, "Missing event start date or end date!"

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
        return False, "Start & end days count cannot differ!"

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



@click.command()
@click.option(
    "--config",
    type=str,
    default="user_settings.json",
    help='"path\\to\\config_file.json". Default: "user_settings.json"'
)
@click.option(
    "--name",
    type=str,
    default="default event title",
    help='event name'
)
@click.option(
    "--descr",
    type=str,
    default="default event description",
    help='event description'
)
@click.option(
    "--start_day",
    type=str,
    default="",
    help='dd/mm/YYYY [dd/mm/YYYY [...]]'
)
@click.option(
    "--start_hr",
    type=str,
    default="",
    help='HH:MM [HH:MM [...]]'
)
@click.option(
    "--end_day",
    type=str,
    default="",
    help='dd/mm/YYYY [dd/mm/YYYY [...]]'
)
@click.option(
    "--end_hr",
    type=str,
    default="",
    help='HH:MM [HH:MM [...]]'
)
@click.option(
    "--loc",
    type=str,
    default="",
    help='event location'
)
@click.option(
    "--cal",
    type=str,
    default="personal",
    help='"calendar-to-use". Default: "personal"'
)
@click.option(
    "--group",
    is_flag=True,
    help='calendar is a group calendar'
)
@click.option(
    "--invite",
    type=str,
    default="",
    help='email(s) to be invited, separated by space'
)
@click.option(
    "--alarm_type",
    type=str,
    default="",
    help='"DISPLAY" or "EMAIL". Alarm to be set on event. Default: none'
)
@click.option(
    "--alarm_format",
    type=str,
    default="",
    help='"h" = hours, "d" = days'
)
@click.option(
    "--alarm_time",
    type=str,
    default="",
    help='time before the event to set an alarm for, in given format'
)
@click.option(
    "--noprompt",
    is_flag=True,
    help='skip user confirmation'
)
@click.option(
    "--noreport",
    is_flag=True,
    help='skip report log copy for developer'
)
@click.option(
    "--noupdate",
    is_flag=True,
    help='skip software updates auto-check'
)


## Main
def main(config, name, descr, start_day, start_hr, end_day, end_hr, loc, cal, group, invite, alarm_type, alarm_format, alarm_time, noprompt, noreport, noupdate):

    global user_settings
    global update_version
    global update_date

    # load user settings from json file
    user_settings = load_user_settings(config)
    if not user_settings:
        err = f"User config file missing: {config}"
        logger.error(err)
        print(err)
        message_box(err, msg_type='error')
        input("Cannot continue. Press enter to exit.")
        return

    # print software header
    print(f"\n{string_header(terminal=True)}")


    # check software updates
    if not noupdate:
        update_version, update_date = check_updates()

        # check dependecies update
        check_dependencies()
        # run self update
        self_update()


    # check command line arguments
    args_ack, err = args_check(start_day, end_day, start_hr, end_hr)
    if not args_ack:
        logger.warning(err)
        print(f"Error: {err}\n")
        print(show_syntax())
        message_box(err, msg_type='warning')
        input("Press enter to exit.")
        return


    # if more than one start & end days/hours are provided, split them and repeat all procedure for each one
    start_day_list = start_day.split()
    end_day_list = end_day.split()
    # split hours
    if start_hr and end_hr:
        start_hr_list = start_hr.split()
        end_hr_list = end_hr.split()

    # if a calendar is given via command line option overrides user settings, if any
    if not cal and 'calendar' in user_settings and user_settings['calendar']:
        cal = user_settings['calendar']
    elif not cal:
        cal = 'personal'

    # if a location is given via command line option overrides user settings, if any
    if not loc and 'location' in user_settings and user_settings['location']:
        loc = user_settings['location']
    
    events_list = []
    # cycle by key over list of event dates and to list one event each
    for i, day in enumerate(start_day_list):

        # build event details
        event_details = {
            'name' : name,
            'description' : descr,
            'calendar' : cal,
            'group' : group,
            'uid' : (f"{str(datetime.now().timestamp())}_{random.randint(100000, 999999)}_{name}@{user_settings['domain']}").replace(" ", "-")
        }
        logger.info(f"Building event details with UID: {event_details['uid']}")

        # add location if any
        if loc:
            event_details['location'] = loc

        # event with fixed hours
        if start_hr and end_hr:
            # hours set to 00:00 equals full day event
            if (start_hr_list[i] == "00:00") and (end_hr_list[i] == "00:00"):
                event_details.update( { 'start' : datetime.strptime(f"{start_day_list[i]}", "%d/%m/%Y").date() } )
                event_details.update( { 'end' : datetime.strptime(f"{end_day_list[i]}", "%d/%m/%Y").date() + timedelta(days=1) } )
                event_details.update( { 'fullday' : True } )
                logger.info(f"Full day event, all-0 hours")
            else:
            # set fixed hours
                event_details.update( { 'start' : datetime.strptime(f"{start_day_list[i]} {start_hr_list[i]}", "%d/%m/%Y %H:%M") } )
                event_details.update( { 'end' : datetime.strptime(f"{end_day_list[i]} {end_hr_list[i]}", "%d/%m/%Y %H:%M") } )
                event_details.update( { 'fullday' : False } )
                logger.info(f"Fixed hours event")
        # full day event
        else:
            event_details.update( { 'start' : datetime.strptime(f"{start_day_list[i]}", "%d/%m/%Y").date() } )
            event_details.update( { 'end' : datetime.strptime(f"{end_day_list[i]}", "%d/%m/%Y").date() + timedelta(days=1) } )
            event_details.update( { 'fullday' : True } )
            logger.info(f"Full day event")
        
        # add invitees, can be 1 or more separated by a space
        if invite:
            event_details.update( { 'invite' : invite } )
            logger.info(f"Invites requested for: {invite}")

        # set alarm - all 3 parameters must be given otherwise none is set
        if alarm_type and alarm_format and alarm_time:
            alarm_type = alarm_type.upper()
            alarm_format = alarm_format.upper()
            if ( alarm_type == 'DISPLAY' or alarm_type == 'EMAIL') and ( alarm_format == 'H' or alarm_format == 'D' ):
                event_details.update( { 'alarm_type' : alarm_type } )
                event_details.update( { 'alarm_format' : alarm_format } )
                event_details.update( { 'alarm_time' : alarm_time } )
                logger.info(f"Alarm requested: {alarm_type}, {alarm_format}, {alarm_time}")

        # append event to list
        events_list.append(event_details)


    # print events summary
    logger.info(f"print events summary")
    output_tk = ''
    print(f"\nThe following {len(start_day_list)} event(s) will be added:\n")

    # cycle over events list
    for j, event_n in enumerate(events_list):

        string_output = (f"Event {j+1}/{len(start_day_list)}\n"
              f"-----------\n"
              f"NAME:           {event_n['name']}\n"
              f"DESCRIPTION:    {event_n['description']}\n"
              f"\nSTART DATE:     {datetime.strftime(event_n['start'], '%d/%m/%Y %H:%M:%S')}\n"
              f"END DATE:       {datetime.strftime(event_n['end'], '%d/%m/%Y %H:%M:%S')}\n"
              f"LOCATION:       {event_n['location'] if 'location' in event_n else 'None'}\n"
              f"CALENDAR:       {event_n['calendar']}")
        if invite:
            string_output += f"INVITEE:        {event_n['invite']}"
        if 'alarm_type' in event_n:
            string_output += f"ALARM:          {event_n['alarm_type']}, {event_n['alarm_time']}{event_n['alarm_format']} before"
        string_output += "\n----------------------------------------------\n"

        # append and print
        output_tk += string_output
        print(string_output)
    
    # skip user confirmation if enabled with --noprompt
    if noprompt:
        logger.info(f"Proceed creating events")
        create_events(events_list)
    else:
        logger.info(f"Wait for user prompt to proceed")
        confirm_box(output_tk, events_list)
        #input("Press enter to confirm.")


    # send log report
    if not noreport:
        send_report()


if __name__ == '__main__':
    main()
