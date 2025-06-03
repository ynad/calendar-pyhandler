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
# calendar-pyCLIent.py
2025-06-03

CLI client to handle operations on calendars: currenlty supports CalDav (WebDAV) and Microsoft Graph API (Microsoft 365).
Event details and other options are via command line arguments. See --help for more.
User settings must be provided in a JSON file. See config examples for CalDav and Graph.

See README.me for full details.
"""

###################################################################################################
# APP SETTINGS - DO NOT EDIT
VERSION_NUM = "0.7.1"
DEV_EMAIL = "info@danielevercelli.it"
DEV_TAG = "ynad"
PROD_REPO = "calendar-pyhandler"
PROD_NAME = "calendar-pyCLIent"
PROD_URL = "github.com/ynad/calendar-pyhandler"
logging_file = "debug.log"
ics_file = "tmp_calendar-pyhandler-event.ics"
###################################################################################################



import sys
import os 
import shutil
import logging
import json
import click
import requests
import random
import signal
import regex as re
from datetime import datetime, timedelta
import zipfile
import tempfile
from packaging import version  # Use packaging.version for semver parsing
import subprocess

# GUI libs
import tkinter as tk
from tkinter import scrolledtext, messagebox

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
# ics tmp file
ics_file=f"{os.path.dirname(__file__)}/{ics_file}"



def message_box(message: str, msg_type: str = 'info') -> None:
    window = tk.Tk()
    window.wm_withdraw()

    if msg_type.lower() == 'error':
        messagebox.showerror(title=f"Error - {string_header(short=True)}", message=message)
    elif msg_type.lower() == 'warning':
        messagebox.showwarning(title=f"Warning - {string_header(short=True)}", message=message)
    else:
        messagebox.showinfo(title=f"Information - {string_header(short=True)}", message=message)

    window.destroy()


def ask_yes_no_gui(message: str, title: str = 'Info', icon: str = 'info', default: str = 'yes'):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    result = messagebox.askyesno(title, message, icon=icon, default=default)
    root.destroy()
    return result


def confirm_events(output_tk: str, events_list: list) -> None:
    # Create main window
    root = tk.Tk()
    root.title(string_header(terminal=False))
    root.geometry("700x400")

    # Label for instruction or title
    label = tk.Label(root, text=f"\nI seguenti eventi saranno creati:\n")
    label.pack(pady=5)

    # ScrolledText widget for displaying the text
    text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=16)
    text.pack(pady=1)

    # insert output text
    text.insert(tk.END, output_tk)
    text.configure(state='disabled')  # Make the text read-only

    # Button frame to organize Confirm and Cancel buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Confirm and Cancel buttons
    confirm_button = tk.Button(button_frame, text="     OK     ", command=lambda: [create_events(events_list), root.destroy(), root.quit()])
    cancel_button = tk.Button(button_frame, text="  Cancel  ", command=lambda: [print('Aborted'), root.destroy(), root.quit()])

    confirm_button.pack(side=tk.LEFT, padx=10)
    cancel_button.pack(side=tk.RIGHT, padx=10)

    # Run the application
    root.mainloop()


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
            "   [--cal \"calendar-name-or-ID\". Default: \"personal\"]\n"
            "   [--group : bool flag to set this as a group calendar. Default: False]\n"
            "   [--invite \"user1@mail.org user2@mail.net\" : email(s) to be invited, separated by space]\n"
            "\nAlarm settings, all 3 parameters must be set or none is considered:\n"
            "   [--alarm_type : \"DISPLAY\" or \"EMAIL\". Alarm to be set on event. Default: none]\n"
            "   [--alarm_format : \"H\" = hours, \"D\" = days]\n"
            "   [--alarm_time : time before the event to set an alarm for. Format HH:MM for \"H\", or N > 0 for \"D\"]\n"
            "\nApp behavior settings:\n"
            "   [--config \"path\\to\\config-file.json\". Default: \"user_settings.json\"]\n"
            "   [--noprompt : bool, skip user confirmation]\n"
            "   [--noreport : bool, skip report log copy for developer]\n"
            "   [--noupdate : bool, skip software updates self-check]\n"
    )


# load user settings from file
def load_user_settings(user_config: str) -> dict:
    logger.info(f"Calendar pyCLIent - v{VERSION_NUM}")
    if os.path.exists(user_config):
        logger.info("Loading user settings JSON from file: " + str(user_config))
        with open(user_config, 'r') as f:
            user_settings = json.load(f)

        # assert mandatory settings keys
        assert 'mode' in user_settings, "invalid user settings, 'mode' key missing"
        assert 'domain' in user_settings, "invalid user settings, 'domain' key missing"
        assert 'username' in user_settings, "invalid user settings, 'username' key missing"

        logger.info(f"Running instance for: {user_settings['domain']}, user: {user_settings['username']}, calendar: {user_settings['calendar'] if 'calendar' in user_settings else 'None'}, mode: {user_settings['mode']}")
        return user_settings
    else:
        logger.error(f"User settings JSON missing: {user_config}")
        return None


# send report of usage to developer
def report_copy(user_settings: dict) -> None:
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


def prompt_yes_no(message: str, default: bool = False) -> bool:
    prompt = " [Y/n]: " if default else " [y/N]: "
    while True:
        reply = input(message + prompt).strip().lower()
        if not reply:
            return default
        if reply in ('y', 'yes', 'Y', 'Yes', 'YES'):
            return True
        if reply in ('n', 'no', 'N', 'No', 'NO'):
            return False


def get_latest_release() -> tuple[str, list]:
    url = f"https://api.github.com/repos/{DEV_TAG}/{PROD_REPO}/releases/latest"
    response = requests.get(url)

    if response.status_code != 200:
        msg = f"Failed to fetch latest release info: {response.status_code} {response.reason}"
        logger.warning(msg)
        print(msg)
        return msg, None

    try:
        data = response.json()
    except Exception as exc:
        msg = f"Exception decoding JSON release info response: {str(exc)}"
        logger.warning(f"Exception decoding JSON release info response: {repr(exc)}")
        print(msg)
        return msg, None

    logger.info(f"Latest release info fetched: {data['tag_name']}, published: {data['published_at']}")
    return data['tag_name'], data['assets']


def is_newer_version(latest_tag: str, current_tag: str) -> bool:
    try:
        return version.parse(latest_tag.lstrip('v')) > version.parse(current_tag.lstrip('v'))
    except Exception as exc:
        logger.warning(f"Exception checking package version: {repr(exc)}")
        return False


def download_zip_asset(assets: list, output_path: str) -> tuple[bool, str]:
    for asset in assets:
        if asset['name'].endswith(".zip"):
            url = asset["browser_download_url"]

            logger.info(f"Downloading update from {url}...")
            print(f"Downloading update from {url}...")

            with requests.get(url, stream=True) as r:
                r.raise_for_status()
                with open(output_path, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            return True, None

    msg = "No ZIP asset found in release, cannot update"
    logger.warning(msg)
    print(msg)
    return False, msg


def unzip_overwrite(zip_path: str, extract_to: str):
    logger.info(f"Extracting {zip_path} to {extract_to}...")
    print(f"Extracting {zip_path} to {extract_to}...")

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for member in zip_ref.namelist():
            zip_ref.extract(member, extract_to)


def update_requirements_if_needed(zip_path: str, temp_extract_dir: str):
    logger.info("Checking for updated Python requirements...")
    print("Checking for updated Python requirements...")

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_extract_dir)

    new_reqs_path = os.path.join(temp_extract_dir, "requirements.txt")
    if not os.path.exists(new_reqs_path):
        logger.info("No requirements.txt found in update")
        print("No requirements.txt found in update")
        return

    #if prompt_yes_no("Do you want to update dependencies? (ATTENTION: if skipped the software might not work after the update)", default=False):
    if ask_yes_no_gui("Do you want to update software dependencies?\n\nWARNING: Skipping this step may cause the software to stop working after the update", title="Dependencies update", icon='question'):
        logger.info("Installing updated dependencies...")
        print("Installing updated dependencies...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "--upgrade", "-r", new_reqs_path])
        except subprocess.CalledProcessError as exc:
            logger.error(f"Failed to update requirements: {e}")
            print(f"Failed to update requirements: {e}")
    else:
        logger.warning("Dependency update skipped")
        print("Dependency update skipped")


def restart_app():
    logger.warning("Restarting app ...")
    print("Restarting app...")
    os.execv(sys.executable, [sys.executable] + sys.argv)
    #os.execl(sys.executable, 'python', __file__, *sys.argv[1:])


def check_and_update():
    latest_tag, assets = get_latest_release()
    if not assets:
        message_box(latest_tag, msg_type='warning')
        return

    if not is_newer_version(latest_tag, VERSION_NUM):
        logger.info(f"No update needed. Current version {VERSION_NUM} is up to date")
        #print(f"No update needed. Current version {VERSION_NUM} is up to date")
        return

    logger.warning(f"New version {latest_tag} is available (current: {VERSION_NUM})")
    print(f"New version {latest_tag} is available (current: {VERSION_NUM})")
    msg = f"New version {latest_tag} is available (current: {VERSION_NUM})\n\nDo you want to download and install it?"

    #if not prompt_yes_no("Do you want to download and install it?", default=True):
    if not ask_yes_no_gui(msg, title="Software update", icon='info'):
        logger.warning("Update canceled by user")
        print("Update canceled by user\n")
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "update.zip")
        res, msg = download_zip_asset(assets, zip_path)
        if not res:
            message_box(msg, msg_type='warning')
            return

        update_requirements_if_needed(zip_path, os.path.join(tmpdir, "extracted"))
        unzip_overwrite(zip_path, os.path.dirname(os.path.abspath(__file__)))

    restart_app()


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


# determine user backend mode and create events accordingly
def create_events(events_list: list) -> None:
    try:
        logger.info(f"create_events, mode: {user_settings['mode']}")

        # CalDav - WebDav
        if user_settings['mode'] == 'caldav':
            agent = CaldavAgent(user_settings, ics_file=ics_file)

        # Microsoft Graph REST API
        elif user_settings['mode'] == 'microsoft_graph':
            agent = MGraphAgent(user_settings)

        else:
            msg = f"Invalid client mode: {user_settings['mode']}, cannot continue"
            logger.error(msg)
            message_box(msg, msg_type='error')
            raise RuntimeError(f"Not implemented client mode: {user_settings['mode']}")

        for event_n in events_list:
            res, msg = agent.create_event(event_n)
            if res:
                message_box(msg, msg_type='info')
            else:
                message_box(msg, msg_type='error')

    except Exception as exc:
        logger.error(f"Exception on create_events: {repr(exc)}")
        print(f"Exception on create_events: {repr(exc)}")

        # show an error message
        message_box(f"Exception on create_events: {str(exc)}", msg_type='error')



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
    default="",
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

    # load user settings from json file
    user_settings = load_user_settings(config)
    if not user_settings:
        err = f"User config file missing: {config}"
        logger.error(err)
        print(err)
        message_box(err, msg_type='error')
        #input("Cannot continue. Press enter to exit.")
        #return 10
        sys.exit(10)

    # print software header
    print(f"\n{string_header(terminal=True)}")

    # check software updates
    if not noupdate:
        check_and_update()


    # check command line arguments
    args_ack, err = args_check(start_day, end_day, start_hr, end_hr)
    if not args_ack:
        logger.warning(err)
        print(f"Error: {err}\n")
        print(show_syntax())
        message_box(err, msg_type='warning')
        #input("Press enter to exit.")
        #return 20
        sys.exit(20)


    # if more than one start & end days/hours are provided, split them and repeat all procedure for each one
    start_day_list = start_day.split()
    end_day_list = end_day.split()
    # split hours
    if start_hr and end_hr:
        start_hr_list = start_hr.split()
        end_hr_list = end_hr.split()

    # if a calendar is given via command line option overrides user settings, if any
    if not cal and 'calendar' in user_settings and len(user_settings['calendar']) > 1:
        cal = user_settings['calendar']
        logger.info(f"Calendar set: {cal}")
    elif not cal:
        cal = 'personal'
        logger.info(f"Calendar default: {cal}")
    

    # set as group calendar if in user settings
    if not group and 'group' in user_settings and user_settings['group'] == True:
        group = True
    # else 'group' stays as set by cmd line option

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
                
                # check HH:MM format
                if alarm_format == 'H':
                    pattern_hour = r'^([0-9]|1[0-9]|2[0-3]):([0-9]|[0-5][0-9])$'
                    if not re.match(pattern_hour, alarm_time):
                        msg = f"Invalid time for alarm_format 'H': 'HH:MM'"
                        logger.warning(msg)
                        message_box(msg, msg_type='warning')
                        raise ValueError(msg)

                # check positive integer for days format
                elif alarm_format == 'D':
                    try:
                        alarm_time = int(alarm_time)
                        assert alarm_time > 0

                    except Exception as exc:
                        msg = f"Invalid time for alarm_format 'D': Integer > 0"
                        logger.warning(msg)
                        message_box(msg, msg_type='warning')
                        raise ValueError(msg)

                event_details.update( { 'alarm_type' : alarm_type } )
                event_details.update( { 'alarm_format' : alarm_format } )
                event_details.update( { 'alarm_time' : alarm_time } )
                logger.info(f"Alarm requested: {alarm_type}, {alarm_format}, {alarm_time}")

            else:
                msg = f"Invalid alarm parameters:\n\n'alarm_type': 'DISPLAY' or 'EMAIL'\n'alarm_format': 'D' or 'H'"
                logger.warning(msg)
                message_box(msg, msg_type='warning')
                raise ValueError(msg)

        # append event to list
        events_list.append(event_details)


    # print events summary
    logger.info(f"print events summary")
    output_tk = ''
    print(f"\nI seguenti ({len(start_day_list)}) eventi saranno creati:\n")

    # cycle over events list
    for j, event_n in enumerate(events_list):

        string_output = (f"Evento {j+1}/{len(start_day_list)}\n"
              f"-----------\n"
              f"NOME:           {event_n['name']}\n"
              f"DESCRIZIONE:    {event_n['description']}\n\n"
              f"DATA INIZIO:    {datetime.strftime(event_n['start'], '%d/%m/%Y %H:%M:%S')}\n"
              f"DATA FINE:      {datetime.strftime(event_n['end'], '%d/%m/%Y %H:%M:%S')}\n"
              f"LUOGO:          {event_n['location'] if 'location' in event_n else 'None'}\n"
              f"CALENDARIO:     {event_n['calendar']}")
        if invite:
            string_output += f"\nINVITATI:       {event_n['invite']}"
        if 'alarm_type' in event_n:
            string_output += f"\nREMINDER:       {event_n['alarm_type']}, {event_n['alarm_time']}{event_n['alarm_format']} prima"
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
        confirm_events(output_tk, events_list)
        #input("Press enter to confirm")


    # send log report
    if not noreport:
        report_copy(user_settings)

    # return success
    return 0


if __name__ == '__main__':
    main()

