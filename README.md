# calendar-pyhandler

Middleware tools to handle operations on calendars: currenlty support Microsoft Graph API (Microsoft 365) and CalDAV (WebDAV).
Aimed to be used as interface for other apps, to easily interact with calendars and automate jobs.

Uses some basic GUI for diplaying operations output.

## User settings
User settings must be formatted as JSON data as follows. By default the app searches for `user_settings.json`. You may specify a different one with `--config` option.

See some examples in `config/` for Graph and CalDAV.

### Microsoft Graph (365):
```
{
    "mode" : "microsoft_graph",
    "azure_client_id" : "secret-azure-client-id",
    "azure_tenant_id" : "secret-azure-tenant-id",
    "domain" : "cloud.domain.com",
    "username": "user.name",
    "calendar" : "personal",
    "_calendar" : "<calendar-or-group-id>",
    "_group" : false,
    "organizer_name" : "User Name",
    "organizer_role" : "Very Specialist",
    "organizer_email" : "user.name@domain.com",
    "_location" : "My Conference Room",
    "_report" : "/tmp/report"
}
```

### CalDAV / WebDAV
```
{
    "mode" : "caldav",
    "domain" : "cloud.domain.com",
    "server" : "https://cloud.domain.com/remote.php/dav/calendars",
    "username": "user.name",
    "password": "super-secret-password",
    "calendar" : "personal",
    "organizer_name" : "User Name",
    "organizer_role" : "Very Specialist",
    "organizer_email" : "user.name@domain.com",
    "_location" : "My Conference Room",
    "_report" : "/tmp/report"
}
```

## Usage
Event details and all other options are set via command line options. See `--help` for more.

Mind the required quotes for each string argument.

Optional arguments are enclosed by [n]; if missing, field is ignored or default values are used.
```
$ python3 calendar-pyCLIent.py
Usage: calendar-pyCLIent.py [OPTIONS]

Syntax:

Event and calendar settings:
    --name "event name"
    --descr "event description"
    --start_day dd/mm/YYYY [dd/mm/YYYY [...]]
    --end_day dd/mm/YYYY [dd/mm/YYYY [...]]
   [--start_hr HH:MM [HH:MM [...]]]
   [--end_hr HH:MM [HH:MM [...]]]
   [--loc "event location"]
   [--cal "calendar-name-or-ID". Default: "personal"]
   [--group : bool flag to set this as a group calendar. Default: False]
   [--invite "user1@mail.org user2@mail.net" : email(s) to be invited, separated by space]

Alarm settings, all 3 parameters must be set or none is considered:
   [--alarm_type : "DISPLAY" or "EMAIL". Alarm to be set on event. Default: none]
   [--alarm_format : "h" = hours, "d" = days]
   [--alarm_time : time before the event to set an alarm for. Format HH:MM for "H", or N > 0 for "D"]

App behavior settings:
   [--config "path\to\config-file.json". Default: "user_settings.json"]
   [--noprompt : bool, skip user confirmation]
   [--noreport : bool, skip report log copy for developer]
   [--noupdate : bool, skip software updates auto-check]
```

## Requirements
- Python >= 3.10
- optional - python venv: set up with `python -m venv .venv`
- pip requirements, listed in `requirements.txt`

For a guided setup of venv and pip requirements use `setup.sh` (Linux) or `setup.bat` (Windows). Or else manually install pip requirements with `pip install -r requirements.txt`.

## License
Released under GPL-3.0 license.

## Disclaimer
Project under early development and tailored for specific usage. Might be expanded or left as is in the future.

Feel free to share any feedback, will be appreciated.
