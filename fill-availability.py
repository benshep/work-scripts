import pandas
from outlook import get_appointments_in_range

dates = pandas.bdate_range()
for appt in get_appointments_in_range(0, 30):
    pass