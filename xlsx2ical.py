import xlrd
from icalendar import Calendar, Event
from re import search
from datetime import datetime, date, time

import argparse

def readXLcalFiles (file,month, year):

    workbook = xlrd.open_workbook(file)

    worksheet = workbook.sheet_by_index(0)       # let's assume it's always in the first sheet
    num_rows = worksheet.nrows - 1
    num_cols = worksheet.ncols - 1
    rowIndex = -1
    inCalendarSection = True
    sched = {}
    dates = []
    while rowIndex < num_rows and inCalendarSection:
        people = []
	rowIndex += 1
	row = worksheet.row(rowIndex)
        rowIsNotEmpty = sum([worksheet.cell_type(rowIndex, m) for m in range (num_cols+1)])
        if rowIsNotEmpty:
            colIndex = -1
            while colIndex < num_cols:
                colIndex += 1
                cell_value = worksheet.cell_value(rowIndex, colIndex)
                if cell_value == 'Residents':    # let's assume the calendar always ends with a
                    inCalendarSection = False    # list of residents
            if max([worksheet.cell_type(rowIndex, m) for m in range (num_cols+1)]) == 2:  # some numbers, must be dates (obv.)
                dates = [worksheet.cell_value(rowIndex, m) for m in range (num_cols+1)]
            elif max([worksheet.cell_type(rowIndex, m) for m in range (num_cols+1)]) == 1: # all text, must be names (obv.)
                people = [worksheet.cell_value(rowIndex, m) for m in range (num_cols+1)]

                for m in zip(dates,people):
                    if m[0]:  # skip empty cells
                        schedDate = "{0}-{1}-{2}".format(year,month,int(m[0])) # format the date,
                        schedDate = datetime.strptime(schedDate, "%Y-%m-%d")   # make it a Python date object.
                        if schedDate in sched:                                 # don't overwrite existing data,
                            if m[1] and not search('Residents',m[1]):       # if there's data and it's not 'Residents'
                                sched[schedDate] = sched[schedDate] +"/"+ m[1] # added it to existing data with a "/"
                        else:                                                  # or else there is no existing data
                            sched[schedDate] = m[1]                            # and what we have is it.

    return(sched)


def getCallsPerPerson(sched):
    person = {}
    for m in sched:
        p = sched[m]
        if (search('/',p)):     # for shared calls, we joined them with a '/', now
            p = p.split('/')[1] # split them on the '/' and take the 2nd field
            p = p.split()[0]    # and split that on a space and take the 1st field

        if p in person:
            person[p] += 1
        else:
            person[p] = 1

    return(person)


def writeICS(file,message,sched):

    ical = Calendar()

    for m in sorted(sched):
        event=Event()
        if search(args.person,sched[m]):   # in case we are filtering, only work on 'args.person'
            event.add('summary',message+sched[m])  # add a summary
            event.add('dtstart',m)                 # add a start (this makes a midnight event, though)
            ical.add_component(event)              # put the event into the calendar

    f = open(file,"wb")
    f.write(ical.to_ical(ical))   # write out our calendar to a nicely formatted .ics file
    f.close


if __name__ == '__main__':

    year = datetime.now().year  # to use as the default year

    parser = argparse.ArgumentParser()
    parser.add_argument ("-y", "--year", type=int, default=year,
                         help="set the year for the calendar; defaults to this year")
    parser.add_argument ("-m", "--month", type=int, required=True,
                         help="set the month being parsed")
    parser.add_argument ("-f", "--file", required=True,
                         help="filename of xlsx file to be read")
    parser.add_argument ("-p", "--person", default='',
                         help="filter results to only this person")
    parser.add_argument ("-c", "--calendarfile",
                         help="filename of calendar (ics) file to write")
    parser.add_argument ("-s", "--summary", type=int, choices=[0,1],
		     nargs='?', const=1,
		     help="print summary stats for number of calls")
    parser.add_argument ("-t", "--text", default='',
                         help="extra text to prepent to the iCal event")
    args = parser.parse_args()

    sched = readXLcalFiles(args.file,args.month, args.year)

    for m in sorted(sched):
        if search(args.person,sched[m]):
            print m.strftime('%a'),str(m).split()[0],sched[m]

    # print the summary, if it was requested
    if args.summary:
        callsPerPerson = getCallsPerPerson(sched)
        print "-"*50
        for m in sorted(callsPerPerson):
            print "{0:>9s}: {1:2d}".format(m,callsPerPerson[m])

    # write out an ICS file if we have a filename
    if args.calendarfile:
        calfilename = args.calendarfile
        if calfilename.split('.')[-1] != 'ics':  # if the filename doesn't end in '.ics'
            calfilename = calfilename+'.ics'     # then it does now.
        print "Writing calendar file to",calfilename  # verbosity! to the screen!
        writeICS(calfilename,args.text,sched)
