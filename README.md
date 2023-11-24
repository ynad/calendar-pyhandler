# caldav-py-handler
A small webdav-caldav client to handle caldav sync from command line.

Aimed to be used as interface for other appliances willing to link with a caldav reality.

## User settings
User settings data must be formatted in JSON as follow, by default the app searches for it in file `user_settings.json`:
```
{
    "domain" : "cloud.domain.com",
    "server" : "https://cloud.domain.com/remote.php/dav/calendars",
    "username": "jane.doe",
    "password": "secret-app-password",
    "calendar_default" : "personal",
    "organizer_name" : "Jane Doe",
    "organizer_role" : "IT",
    "organizer_email" : "info@example.com",
    "location_default" : "Main Office",
    "report" : "path/to/reports-folder"
}
```
`USER SETTINGS` section in the code might need to be adjusted to your working path.
Logging level defaults to DEBUG, to log file `debug.log`.

## Syntax
Mind the required quotes for each string argument.

Optional arguments are enclosed by [n]; if missing, field is ignored or default values are used.
```
python caldav-ics-client.py

Missing arguments! Syntax:
caldav-ics-client.py
    --name "event name"
    --descr "event description"
    --start_day dd/mm/YYYY
    --end_day dd/mm/YYYY
   [--start_hr HH:MM:SS]
   [--end_hr HH:MM:SS]
   [--loc "event location"]
   [--cal "calendar to be used"]
   [--invite "email(s) to be invited, separated by space"]
   [--alarm_type : alarm to be set on event: "DISPLAY" or "EMAIL". Default: none]
   [--alarm_format : "h" = hours, "d" = days]
   [--alarm_time : time before the event to set an alarm for]
   [--config "path/to/config-file.json. Default: "user_settings.json"]
   [--prompt "y/n" : wait or skip user confirmation. Default: y]
   [--report "y/n" : save report log for developer. Default: y]\n"
   [--update "y/n" : Auto-check software updates. Default: y]\n"
```

## Requirements
Required python libraries are listed in file `requirements.txt`.
Install them with pip:
```
pip install -r requirements.txt
```


## License
Released under GPL-3.0 license.

## Disclaimer
Project under early development and tailored for specific usage. Might be expanded or left as is in the future.

Feel free to share any feedback, will be appreciated.
