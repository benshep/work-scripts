import re
import outlook
import datefinder
from datetime import datetime

outlook_app = outlook.get_outlook()
namespace = outlook_app.GetNamespace('MAPI')
inbox = namespace.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort('[Received]', Descending=True)

for _ in range(10):
    message = messages.GetNext()
    print(message.Subject, message.FlagRequest)
    if not message.UnRead and message.FlagRequest:
        break
# print(message.Subject, message.LastModificationTime, message.UnRead)

best = None
longest_source = ''
now = datetime.now()
next_year_months = range(1, (now.month + 7) - 12)
for date, source in datefinder.find_dates(message.Body, source=True, first='day'):
    print(f'{date}, "{source}"')
    if date.month in next_year_months and date.year == now.year:
        date = date.replace(year=date.year + 1)
    if date < now:
        continue
    if len(source) > len(longest_source):
        longest_source = source
        best = date

print(f'{best}, "{longest_source}"')
calendar = outlook.get_calendar()
event_title = re.sub('^(Re|Fw|Fwd): ', '', message.Subject, flags=re.IGNORECASE)

event = outlook_app.CreateItem(1)  # AppointmentItem
event.Subject = event_title
event.Body = message.Body
event.Start = best
event.AllDayEvent = False
event.Location = 'Zoom' if 'Zoom' in message.Body else 'Teams' if 'Teams' in message.Body else ''
event.Duration = 60  # can do better than this?
event.BusyStatus = 2  # busy
event.Display()

input()