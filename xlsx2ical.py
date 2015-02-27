import xlrd
import re
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
            if max([worksheet.cell_type(rowIndex, m) for m in range (num_cols+1)]) == 2:
                dates = [worksheet.cell_value(rowIndex, m) for m in range (num_cols+1)]
            elif max([worksheet.cell_type(rowIndex, m) for m in range (num_cols+1)]) == 1:
                people = [worksheet.cell_value(rowIndex, m) for m in range (num_cols+1)]

                for m in zip(dates,people):
                    if m[0]:  # skip empty cells
                        schedDate = "{0}-{1}-{2}".format(year,month,int(m[0])) # format the date
                        schedDate = datetime.strptime(schedDate, "%Y-%m-%d")   # make it a Python date object
                        if schedDate in sched:                                 # don't overwrite existing data
                            if m[1] and not re.search('Residents',m[1]):       # if there's data and it's not 'Residents'
                                sched[schedDate] = sched[schedDate] +"/"+ m[1] # added to existing with a "/"
                        else:                                                  # or else there is no existing data
                            sched[schedDate] = m[1]



    return(sched)


def getCallsPerPerson(sched):
    person = {}
    for m in sched:
        p = sched[m]
        if (re.search('/',p)):
            p = p.split('/')[1]
            p = p.split()[0]

        if p in person:
            person[p] += 1
        else:
            person[p] = 1

    return(person)


if __name__ == '__main__':

    year = datetime.now().year

    parser = argparse.ArgumentParser()
    parser.add_argument ("-y", "--year", type=int, default=year,
                         help="set the year for the calendar; defaults to this year")
    parser.add_argument ("-m", "--month", type=int, required=True,
                         help="set the month being parsed")
    parser.add_argument ("-f", "--file", required=True,
                         help="filename of xlsx file to be read")
    parser.add_argument ("-p", "--person", default='',
                         help="filter results to only this person")
    parser.add_argument ("-s", "--summary", type=int, choices=[0,1],
		     nargs='?', const=1,
		     help="print summary stats for number of calls")
    args = parser.parse_args()

    sched = readXLcalFiles(args.file,args.month, args.year)

    for m in sorted(sched):
        if re.search(args.person,sched[m]):
            print str(m).split()[0],sched[m]
    if (args.summary):
        callsPerPerson = getCallsPerPerson(sched)
        print "-"*50
        for m in sorted(callsPerPerson):
            print "{0:>9s}: {1:2d}".format(m,callsPerPerson[m])