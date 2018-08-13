import time
import datetime
#day month year
s = "23/07/2018"
UnixTime = time.mktime(datetime.datetime.strptime(s,"%d/%m/%Y").timetuple())

print(UnixTime)