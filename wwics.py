#!/usr/bin/env python2.7
from __future__ import print_function
from icalendar import Calendar, Event
from datetime import datetime, timedelta
from time import strptime, strftime, mktime

""" See https://secure.wideworld-sports.me/wws_membership/PrintSchedule.asp?ID=1439
and https://secure.wideworld-sports.me/wws_membership/DivisionReport.asp?ID=1439

This could be improved to go read the schedule from the web page based
on season, team name, etc.  Or just copying-and-pasting the schedule
to a text file and using grep and stuff like this one does """


def writeICS(file, sched):

    location = "Wide World Sports, 2140 Oak Valley Dr, Ann Arbor, MI 48103"
    ical = Calendar()

    for m in sorted(sched):
        event=Event()
        event.add('summary', "Fairies Soccer on "+sched[m]['desc'])
        event.add('dtstart', sched[m]['starttime'])
        event.add('dtend', sched[m]['endtime'])
        event.add('location', location)

        ical.add_component(event)

    # write out our calendar to a nicely formatted .ics file
    f = open(file,"wb")
    f.write(ical.to_ical(ical))
    f.close


if __name__ == '__main__':
    FILE="./ff"  # hopefully there is a file called 'ff' formated like:
    # 1/14/2017	Sat 	1:30 PM	Big Georges Field	Fairies	 	WWSC 8-9 Portugal
    # with tabs separating the items

    with open (FILE, "rb") as f:
        a=f.readlines()

    a = [b.strip('\n') for b in a]

    game = dict()
    for evt in a:
        # break up the line on tabs
        line = evt.split('\t')
        # cope with the date and time
        ds = strftime('%Y-%m-%dT%H:%M',
                      strptime(line[0]+" "+line[2],
                                        "%m/%d/%Y %I:%M %p"))
        # make a datetime object so icalendar.event.add works
        dss = datetime.fromtimestamp(
            mktime(strptime(line[0]+" "+line[2],
                            "%m/%d/%Y %I:%M %p")))
        # make the duration for the end time
        dse = dss + timedelta(hours=1)
        # make the game details (start, end, description)
        game[ds] = dict()
        game[ds]['desc'] = "{}: {} vs {}".format(
            line[3],
            line[4],
            line[6])
        game[ds]['starttime'] = dss
        game[ds]['endtime'] = dse

    # print a nice schedule to the screen
    for d in sorted(game):
        print ("{} -  {}".format(d, game[d]['desc']))

    # write the ICS file for importing into a calendar program
    writeICS("FairiesWinter2017Schedule.ics", game)
