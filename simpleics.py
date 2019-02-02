#!/usr/bin/env python2.7
from __future__ import print_function
from icalendar import Calendar, Event
from datetime import datetime, timedelta, date
from time import strptime, strftime, mktime


def writeICS(file, sched):

    location = ""
    ical = Calendar()

    for m in sorted(sched):
        event = Event()
        event.add("summary", sched[m]["desc"])
        event.add("dtstart", sched[m]["starttime"])
        event.add("location", location)

        ical.add_component(event)

    # write out our calendar to a nicely formatted .ics file
    f = open(file, "wb")
    f.write(ical.to_ical(ical))
    f.close


if __name__ == "__main__":
    FILE = "./ff"  # hopefully there is a file called 'ff' with a list of dates in it
    DESC = "ENTER YOUR DESCRIPTION HERE"
    OUTPUT_FILE = "calendar.ics"

    with open(FILE, "rb") as f:
        a = f.readlines()

    a = [b.strip("\n") for b in a]

    game = dict()
    for evt in a:
        if evt is "":
            continue
        ds = strftime("%m/%d/%Y", strptime(evt, "%m/%d/%Y"))
        dss = date.fromtimestamp(mktime(strptime(evt, "%m/%d/%Y")))

        game[ds] = dict()
        game[ds]["desc"] = DESC
        game[ds]["starttime"] = dss

    # print a nice schedule to the screen
    for d in sorted(game):
        print("{} -  {}".format(d, game[d]["desc"]))

    # write the ICS file for importing into a calendar program
    writeICS(OUTPUT_FILE, game)
