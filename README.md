# loginmon

Perl service to monitor and record Win32 logins.

Saves to local sqlite db and can synchronise with remote database (currently targetting PostgreSQL hosted [MRBS](http://mrbs.sourceforge.net/) db.)

# Requirements

* Win32::Daemon
* DBI
 * DBD::Sqlite
 * DBD::Pg
* YAML

(tested with Strawberry perl)

# Installation

1. Put it in a folder
2. Edit the example yaml config
3. Run with -inst flag
4. Start the service (or reboot)

# Usage
```
NAME
    loginmon.pl - A windows service which records who is logged in

SYNOPSIS
    loginmon.pl path\loginmon.yaml [-option]

    Options:

    -inst   Install the service, initialise the local database

    -uninst Uninstall the service

    -initdb Initialise the local database (deletes any existing version!)

    -syncdb Test the database sync
```

