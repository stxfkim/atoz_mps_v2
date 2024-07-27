import pandas as pd
from datetime import datetime

# create datetime objects with a fixed date (in this case, January 1, 2000)
time1 = pd.Timestamp('08:00:00')
time2 = pd.Timestamp('04:20:12')

# calculate the time difference
#delta = pd.Timedelta(time2 - time1)
delta = time2 - time1
hours = delta.components.hours
minutes = delta.components.minutes
seconds = delta.components.seconds

# print the result
print(delta,hours,minutes,seconds)