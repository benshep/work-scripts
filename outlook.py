import os
from time import sleep
from typing import Protocol, Callable

import win32com.client
import pywintypes
import re
from datetime import date, datetime, timedelta
import pythoncom
from icalendar import Calendar
from folders import hr_info_folder
from enum import IntEnum


class DefaultFolders(IntEnum):
    """Specifies the folder type for a specified folder in Outlook."""
    deleted_items = 3
    """The Deleted Items folder."""

    outbox = 4
    """The Outbox folder."""

    sent_mail = 5
    """The Sent Mail folder."""

    inbox = 6
    """The Inbox folder."""

    calendar = 9
    """The Calendar folder."""

    contacts = 10
    """The Contacts folder."""

    journal = 11
    """The Journal folder."""

    notes = 12
    """The Notes folder."""

    tasks = 13
    """The Tasks folder."""

    drafts = 16
    """The Drafts folder."""

    all_public_folders = 18
    """The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account."""

    conflicts = 19
    """The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account."""

    sync_issues = 20
    """The Sync Issues folder. Only available for an Exchange account."""

    local_failures = 21
    """The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account."""

    server_failures = 22
    """The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account."""

    junk = 23
    """The Junk E-Mail folder."""

    rss_feeds = 25
    """The RSS Feeds folder."""

    to_do = 28
    """The To Do folder."""

    managed_email = 29
    """The top-level folder in the Managed Folders group. Only available for an Exchange account."""

    suggested_contacts = 30
    """The Suggested Contacts folder."""


class ResponseStatus(IntEnum):
    """Indicates the response to a meeting request."""
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.olresponsestatus

    accepted = 3
    """Meeting accepted."""
    declined = 4
    """Meeting declined."""
    none = 0
    """The appointment is a simple appointment and does not require a response."""
    not_responded = 5
    """Recipient has not responded."""
    organized = 1
    """The AppointmentItem is on the Organizer's calendar or the recipient is the Organizer of the meeting."""
    tentative = 2
    """Meeting tentatively accepted."""


class Recipient(Protocol):
    """Represents a user or resource in Outlook, generally a mail or mobile message addressee."""
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.recipient

    Name: str
    """The display name for the Recipient."""
    Address: str
    """The email address of the Recipient."""
    MeetingResponseStatus: ResponseStatus
    """The overall status of the response to the meeting request for the recipient."""
    AddressEntry: object
    Application: object
    AutoResponse: str
    """The text of an automatic response for a Recipient."""
    Class: object
    DisplayType: object
    EntryID: str
    """The unique Entry ID of the object."""
    Index: int
    """The position of the object within the collection."""
    Parent: object
    PropertyAccessor: object
    Resolved: bool
    """Indicates True if the recipient has been validated against the Address Book."""
    Sendable: bool
    """Indicates whether a meeting request can be sent to the Recipient."""
    Session: object
    TrackingStatus: object
    TrackingStatusTime: object
    Type: int

    def Delete(self) -> None:
        """Deletes an object from the collection."""
        pass

    def FreeBusy(self, start: pywintypes.TimeType, min_per_char: int, complete_format: bool = False) -> str:
        """Returns free/busy information for the recipient.
        The default is to return a string representing one month of free/busy information compatible with the
        Microsoft Schedule+ Automation format (that is, the string contains one character for each MinPerChar minute,
        up to one month of information from the specified Start date).
        If the optional argument complete_format is omitted or False, then "free" is indicated by the character 0
        and all other states by the character 1.
        If CompleteFormat is True, then the same length string is returned as defined above,
        but the characters now correspond to the OlBusyStatus constants."""

    def Resolve(self) -> bool:
        """Attempts to resolve a Recipient object against the Address Book.
        Returns true if the object was resolved; otherwise, false."""


class Folder(Protocol):
    """Represents an Outlook folder."""


class SyncObject(Protocol):
    """Represents a Send/Receive group for a user.
    A Send/Receive group lets users configure different synchronization scenarios, selecting which folders and which filters apply."""
    Name: str
    """The display name for the object."""

    def Start(self):
        """Begins synchronizing a user's folders using the specified Send/Receive group."""

    def Stop(self):
        """Immediately ends synchronizing a user's folders using the specified Send/Receive group."""


class NameSpace(Protocol):
    """Represents an abstract root object for any data source."""
    SyncObjects: list[SyncObject]
    """Contains a set of SyncObject objects representing the Send/Receive groups for a user."""

    def GetDefaultFolder(self, folder_type: DefaultFolders) -> Folder:
        """Returns a Folder object that represents the default folder of the requested type for the current profile;
        for example, obtains the default Calendar folder for the user who is currently logged on."""

    def CreateRecipient(self, recipient_name: str) -> Recipient:
        """Creates a Recipient object."""

    def GetSharedDefaultFolder(self, recipient: Recipient, folder_type: DefaultFolders) -> Folder:
        """Returns a Folder object that represents the specified default folder for the specified user."""


class OutlookApplication(Protocol):
    """Represents the entire Microsoft Outlook application."""
    def GetNamespace(self, namespace_type: str) -> NameSpace:
        """Returns a NameSpace object of the specified type."""


class OlBusyStatus(IntEnum):
    busy = 2
    """The user is busy."""
    free = 0
    """The user is available."""
    out_of_office = 3
    """The user is out of office."""
    tentative = 1
    """The user has a tentative appointment scheduled."""
    working_elsewhere = 4
    """The user is working in a location away from the office."""


class AppointmentItem(Protocol):
    """Represents a meeting, a one-time appointment, or a recurring appointment or meeting in the Calendar folder."""
    # https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem

    Subject: str
    """The subject for the Outlook item."""
    Body: str
    """The clear-text body of the Outlook item."""
    AllDayEvent: bool
    """Returns True if the appointment is an all-day event (as opposed to a specified time)."""
    BusyStatus: OlBusyStatus
    """The busy status of the user for the appointment."""
    Start: pywintypes.TimeType
    """The starting date and time for the Outlook item."""
    StartUTC: pywintypes.TimeType
    """The start date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard."""
    End: pywintypes.TimeType
    """The end date and time for the Outlook item."""
    EndUTC: pywintypes.TimeType
    """The end date and time of the appointment expressed in the Coordinated Universal Time (UTC) standard."""
    Duration: int
    """Duration in minutes."""
    Recipients: list[Recipient]
    """All the recipients for the Outlook item."""
    RequiredAttendees: str
    """A semicolon-delimited String of required attendee names for the meeting appointment."""
    OptionalAttendees: str
    """A semicolon-delimited String of optional attendee names for the meeting appointment."""
    Location: str
    """The specific office location (for example, Building 1 Room 1 or Suite 123) for the appointment."""
    ReminderSet: bool
    """Indicates whether a reminder has been set for the appointment."""
    ReminderMinutesBeforeStart: int
    """The number of minutes before the start of the appointment when the reminder should occur."""
    MeetingStatus: int
    """The status of the meeting (e.g., non-meeting, meeting, meeting request, meeting cancellation)."""
    Importance: int
    """The importance level of the appointment (e.g., low, normal, high)."""
    Sensitivity: int
    """The sensitivity level of the appointment (e.g., normal, personal, private, confidential)."""
    Categories: str
    """The categories assigned to the appointment."""
    Class: int
    """OlObjectClass constant indicating the object's class."""
    CreationTime: pywintypes.TimeType
    """The creation time of the appointment."""
    LastModificationTime: pywintypes.TimeType
    """The last modification time of the appointment."""
    RecurrenceState: int
    """The recurrence state of the appointment."""
    ResponseRequested: bool
    """Indicates whether a response is requested for the appointment."""
    SendUsingAccount: str
    """The account used to send the appointment."""
    StartInStartTimeZone: datetime
    """The start time in the start time zone."""
    EndInEndTimeZone: datetime
    """The end time in the end time zone."""
    IsRecurring: bool
    """Indicates whether the appointment is recurring."""
    MeetingWorkspaceURL: str
    """The URL of the meeting workspace."""

    Actions: object
    Application: object
    Attachments: object
    AutoResolvedWinner: object
    BillingInformation: object
    Companies: object
    Conflicts: object
    ConversationID: object
    ConversationIndex: object
    ConversationTopic: object
    DownloadState: object
    EndTimeZone: object
    EntryID: object
    ForceUpdateToAllAttendees: object
    FormDescription: object
    GetInspector: object
    GlobalAppointmentID: object
    InternetCodepage: object
    IsConflict: object
    ItemProperties: object
    MarkForDownload: object
    MessageClass: object
    Mileage: object
    NoAging: object
    Organizer: object
    OutlookInternalVersion: object
    OutlookVersion: object
    Parent: object
    PropertyAccessor: object
    ReminderOverrideDefault: object
    ReminderPlaySound: object
    ReminderSoundFile: object
    ReplyTime: object
    Resources: object
    ResponseStatus: object
    RTFBody: object
    Saved: object
    Session: object
    Size: object
    StartTimeZone: object
    UnRead: object
    UserProperties: object

    def Save(self) -> None:
        """Saves the appointment item."""

    def Delete(self) -> None:
        """Deletes the appointment item."""

    def Send(self) -> None:
        """Sends the appointment item."""

    def Display(self) -> None:
        """Displays the appointment item."""

    def Close(self) -> None:
        """Closes the appointment item."""

    def ForwardAsVcal(self) -> object:
        """Forwards the AppointmentItem as a vCal; virtual calendar item."""


class SyncState(IntEnum):
    sync_started = 1
    """Synchronization started."""
    sync_stopped = 0
    """Synchronization stopped."""


class SyncObjectEventHandler:
    name: str = ''
    """The name of the sync object."""
    syncing: bool = False
    """True if the object is currently syncing, otherwise False."""
    description: str = ''
    """A description of the current state of the sync process."""
    value: int = 0
    """The current value of the synchronization process (such as the number of items synchronized)."""
    max_value: int = 0
    """The maximum that `value` can reach."""
    sync_start_callback: Callable[[str], None] | None = None
    sync_end_callback: Callable[[str], None] | None = None
    progress_callback: Callable[[str, SyncState, str, int, int], None] | None = None
    error_callback: Callable[[str, int, str], None] | None = None

    def set_name(self, name: str) -> None:
        """Set the name of the object being synced."""
        self.name = name

    def set_callbacks(self,
                      sync_start_callback: Callable[[str], None] | None = None,
                      sync_end_callback: Callable[[str], None] | None = None,
                      progress_callback: Callable[[str, SyncState, str, int, int], None] | None = None,
                      error_callback: Callable[[str, int, str], None] | None = None,
                      ):
        """Specify a function to be called when a given event is raised."""
        self.sync_start_callback = sync_start_callback
        self.sync_end_callback = sync_end_callback
        self.progress_callback = progress_callback
        self.error_callback = error_callback

    def OnSyncStart(self) -> None:
        """Event handler for the start of the sync process."""
        print(f"{self.name}: sync started")
        self.syncing = True
        if self.sync_start_callback:
            self.sync_start_callback(self.name)

    def OnSyncEnd(self) -> None:
        """Event handler for the end of the sync process."""
        print(f"{self.name}: sync completed")
        self.syncing = False
        if self.sync_end_callback:
            self.sync_end_callback(self.name)

    def OnProgress(self, state: SyncState, description: str, value: int, max_value: int) -> None:
        """Occurs periodically while Outlook is synchronizing a user's folders using the specified Send/Receive group.
        :param state: A value that identifies the current state of the synchronization process.
        :param description: A textual description of the current state of the synchronization process.
        :param value: Specifies the current value of the synchronization process (such as the number of items synchronized).
        :param max_value: The maximum that `value` can reach. The ratio of `value` to `max_value` represents the percent complete of the synchronization process.
        """
        print(f"{self.name}: synced {value} of {max_value}, {SyncState(state).name}, {description}")
        self.syncing = state == SyncState.sync_started
        self.description = description
        self.value = value
        self.max_value = max_value
        if self.progress_callback:
            self.progress_callback(self.name, state, description, value, max_value)

    def OnError(self, code: int, description: str) -> None:
        """Occurs when Outlook encounters an error while synchronizing a user's folders using the specified Send/Receive group.
        :param code: A unique value that identifies the error.
        :param description: A textual description of the error.
        """
        print(f'{self.name}: error occurred, {code=}, {description=}')
        self.syncing = False
        if self.error_callback:
            self.error_callback(self.name, code, description)


def perform_sync(sync_start_callback: Callable[[str], None] | None = None,
                 sync_end_callback: Callable[[str], None] | None = None,
                 progress_callback: Callable[[str, SyncState, str, int, int], None] | None = None,
                 error_callback: Callable[[str, int, str], None] | None = None,
                 ) -> None:
    """Carries out a sync of all the SyncObjects in Outlook's MAPI namespace.
    If a sync_end_callback is provided, the function will return immediately and sync will continue in the background.
    Otherwise, this function blocks until sync is completed.
    Supply other callbacks to be notified of sync start, progress and errors."""
    namespace = get_outlook().GetNamespace('MAPI')
    sync_objects = namespace.SyncObjects
    event_handlers: list[SyncObjectEventHandler] = []
    for sync_object in sync_objects:
        sync_object_events: SyncObjectEventHandler = win32com.client.WithEvents(sync_object, SyncObjectEventHandler)
        event_handlers.append(sync_object_events)
        sync_object_events.set_name(sync_object.Name)
        sync_object_events.set_callbacks(
            sync_end_callback=sync_end_callback,
            sync_start_callback=sync_start_callback,
            progress_callback=progress_callback,
            error_callback=error_callback
        )
        sync_object.Start()
    if sync_end_callback:
        return
    while all(handler.syncing for handler in event_handlers):
        pythoncom.PumpWaitingMessages()
        sleep(0.1)



date_spec = datetime | date | float | int
"""A date/datetime or relative days from today."""


def get_appointments_in_range(start: date_spec = 0.0,
                              end: date_spec = 30.0,
                              user: str = 'me') -> list[AppointmentItem]:
    """Get a list of appointments in the given range. start and end are relative days from today, or datetimes."""
    today = date.today()
    from_date = today + timedelta(days=start) if isinstance(start, (int, float)) else start - timedelta(days=1)
    to_date = today + timedelta(days=end) if isinstance(end, (int, float)) else end
    date_filter = f"[Start] >= '{datetime_text(from_date)}' AND [Start] <= '{datetime_text(to_date)}'"
    return get_appointments(date_filter, user=user)


def get_appointments(restriction: str, sort_order: str = 'Start', user: str = 'me') -> list[AppointmentItem]:
    """Get a list of calendar appointments with the given Outlook filter."""
    calendar = get_calendar(user)
    appointments = calendar.Items
    appointments.Sort(f"[{sort_order}]")
    appointments.IncludeRecurrences = True
    return appointments.Restrict(restriction)


def get_calendar(user: str = 'me') -> Folder:
    """Return the calendar folder for a given user. If none supplied, default to my calendar."""
    namespace = get_outlook().GetNamespace('MAPI')
    if user == 'me':
        return namespace.GetDefaultFolder(DefaultFolders.calendar)
    recipient = namespace.CreateRecipient(user)
    if not recipient.Resolve():
        raise RuntimeError(f'User "{user}" not found.')
    return namespace.GetSharedDefaultFolder(recipient, DefaultFolders.calendar)


def get_outlook() -> OutlookApplication:
    """Return a reference to the Outlook application."""
    pythoncom.CoInitialize()  # try to combat the "CoInitialize has not been called" error
    return win32com.client.Dispatch('Outlook.Application')


def get_meeting_time(event: AppointmentItem, get_end: bool = False) -> datetime:
    """Return the start or end time of the meeting (in UTC) as a Python datetime."""
    meeting_time = event.EndUTC if get_end else event.StartUTC
    return datetime.fromtimestamp(meeting_time.timestamp())


def happening_now(event: AppointmentItem, hours_ahead: float = 0.5) -> bool:
    """Return True if an event is currently happening or is due to start in the next half-hour."""
    try:
        # print(appointmentItem.Start)
        start_time = get_meeting_time(event)
        end_time = get_meeting_time(event, get_end=True)
        buffer = timedelta(hours=hours_ahead)
        return start_time - buffer <= datetime.now() <= end_time + buffer
    except (OSError, pywintypes.com_error):  # appointments with weird dates!
        return False


def get_current_events(user: str = 'me', hours_ahead: float = 0.5) -> list[AppointmentItem]:
    """Return a list of current events from Outlook, sorted by subject."""
    current_events = filter(lambda event: happening_now(event, hours_ahead),
                            get_appointments_in_range(-7, 1 + hours_ahead / 24, user=user))
    # Sort by start time then subject, so we have a predictable order for events to be returned
    return sorted(current_events, key=lambda event: f'{event.StartUTC} {event.Subject}')


def datetime_text(advance_time: datetime) -> str:
    return advance_time.strftime('%#d/%#m/%Y %H:%M')


def is_my_all_day_event(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is the user's own all day event."""
    return event.AllDayEvent and len(event.Recipients) <= 1  # not sent by someone else


def is_out_of_office(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is an out of office booking."""
    return is_my_all_day_event(event) and event.BusyStatus == OlBusyStatus.out_of_office


def is_annual_leave(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is an annual leave booking."""
    return is_out_of_office(event) and any([event.Subject.lower().endswith('annual leave'), event.Subject == 'Off',
                                            re.search(r'\bAL$', event.Subject)])  # e.g. "ARB AL" but not "INTERNAL"


def is_wfh(event: AppointmentItem) -> bool:
    """Check whether a given Outlook event is a work from home day."""
    return all([is_my_all_day_event(event),
                event.BusyStatus == OlBusyStatus.working_elsewhere,
                re.search(r'\bWork(ing)? from home$', event.Subject) or re.search(r'\bWFH$', event.Subject)])


def get_away_dates(start: date_spec = -30, end: date_spec = 90,
                   user: str = 'me', look_for: callable = is_out_of_office) -> set[date]:
    """Return a set of the days in the given range that the user is away from the office.
    The look_for parameter can be:
     - is_my_all_day_event: look for any all day event with only the user as an attendee
     - is_out_of_office (default): as above, but set to out of office
     - is_annual_leave: as above, but subject is "Annual Leave" or "AL"
     - is_wfh: set to out of office, and subject is "Work(ing) from home" or "WFH"
     """
    events = filter(look_for, get_appointments_in_range(start, end, user=user))
    # Need to subtract a day here since the end time of an all-day event is 00:00 on the next day
    try:
        away_list = [get_date_list(event.Start.date(),
                                   event.End.date() - timedelta(days=1))
                     for event in events]
        return to_set(away_list)
    except pywintypes.com_error:
        print(f"Warning: couldn't fetch away dates for {user}")
        return set()


def get_date_list(start: date,
                  end: date | None = None,
                  count: int = 0) -> list[date]:
    """Return a list of business dates (i.e. Mon-Fri) in a given date range, inclusive.
    Specify count instead of end to return a fixed number of dates."""
    if end is None and count == 0:  # neither optional argument specified - assume one day
        count = 1
    date_list = []
    while True:
        weekday = start.weekday()
        if weekday < 5:  # Mon-Fri
            date_list.append(start)
            if 0 < count == len(date_list):
                break
        start += timedelta(days=1)
        if end and start > end:
            break
    return date_list


def to_set(date_lists: list[list]) -> set:
    """Convert a list of date lists to a flat set."""
    return set(sum(date_lists, []))


def list_meetings():
    event_list = get_appointments_in_range(start=date(2022, 4, 1), end=date(2023, 3, 31))
    for event in event_list:
        print(event.Duration, event.AllDayEvent, event.Subject, OlBusyStatus(event.BusyStatus), sep='\t')
    print(len(event_list))


def get_dl_ral_holidays() -> set[date]:
    """Return a list of holiday dates for DL/RAL. Fetches ICS file from STFC HR info folder (synced via OneDrive)."""
    filename = os.path.join(hr_info_folder, f'DL_RAL_Site_Holidays_{datetime.now().year}.ics')
    calendar = Calendar.from_ical(open(filename, encoding='utf-8').read())
    return {event.decoded('dtstart') for event in calendar.walk('VEVENT')}


if __name__ == '__main__':
    # away_dates = sorted(
    #     list(get_away_dates(datetime.date(2024, 2, 12), 0, look_for=is_annual_leave)))
    # print(len(away_dates))
    # print(*away_dates, sep='\n')
    print(list_meetings())
