import grabber
import datetime
import calendar
day = []

def grabberFunction(year, month):

    a = calendar.monthcalendar(year, month)
    dayofmonth = []

    for i in a:
        dayofmonth.extend(i)
    dayofmonth = list(filter(lambda num: num != 0, dayofmonth))

    if month == datetime.datetime.now().month and year == datetime.datetime.now().year:
        dayofmonth = dayofmonth[0:datetime.datetime.now().day-1]
        '''YYYY-MM-DD'''

    if month < 10:
        inputmonth = "0" + str(month)
    else:
        inputmonth = str(month)
    for i in dayofmonth:
        day.append(i)
        if i < 10:
            dayofmonth[i-1] = str(year)+"-"+inputmonth+"-"+"0"+str(i)
        else:
            dayofmonth[i-1] = str(year)+"-"+inputmonth+"-"+str(i)
    grabber.loginAccess()
    grabber.downloader(dayofmonth)
