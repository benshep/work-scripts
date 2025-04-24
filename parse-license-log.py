import os
import re
from enum import Enum
from typing import Any, Protocol

import dateutil.parser
from collections import defaultdict
import datetime
import operator
import win32com.client
import sys
from folders import docs_folder


class ADSNameType(Enum):
    """Specifies the format of the name used to identify an object."""
    type_1779 = 1
    """Distinguished Name (e.g., CN=Jeff Smith,OU=Sales,DC=Fabrikam,DC=COM)."""

    canonical = 2
    """Canonical Name (e.g., fabrikam.com/Users/Jeff Smith)."""

    nt4 = 3
    """NT 4.0 Account Name (e.g., FABRIKAM\JeffSmith)."""

    display = 4
    """Display Name (e.g., Jeff Smith)."""

    domain_simple = 5
    """Simple Domain Name (e.g., JeffSmith)."""

    enterprise_simple = 6
    """Enterprise Simple Name (e.g., JeffSmith)."""

    guid = 7
    """Globally Unique Identifier (GUID) (e.g., {95ee9fff-3436-11d1-b2b0-d15ae3ac8436})."""

    unknown = 8
    """Unknown Name Type."""

    user_principal_name = 9
    """User Principal Name (e.g., jeffsmith@fabrikam.com)."""

    canonical_ex = 10
    """Extended Canonical Name (e.g., fabrikam.com/Users Jeff Smith)."""

    service_principal_name = 11
    """Service Principal Name (e.g., www/www.fabrikam.com@fabrikam.com)."""

    sid_or_sid_history_name = 12
    """Security Identifier (SID) or SID History Name."""

    dns_domain = 13
    """DNS Domain Name (e.g., fabrikam.com)."""


class NameTranslate(Protocol):
    def Set(self, set_type: ADSNameType, ads_path: str) -> None:
        """Directs the directory service to set up a specified object for name translation."""

    def Get(self, format_type: ADSNameType):
        """Retrieves the name of a directory object in the specified format. The distinguished name must have been
        set in the appropriate format by the NameTranslate.Set method."""


name_resolver = win32com.client.Dispatch(dispatch='NameTranslate')


class User():
    def __init__(self):
        self.checkout_time = None
        self.checkouts = []
        self.denied = 0

    def check_out(self, timestamp):
        if self.checkout_time is None:
            self.checkout_time = timestamp
    
    def check_in(self, timestamp):
        if not self.checkout_time is None:
            self.checkouts.append((self.checkout_time, timestamp))
            self.checkout_time = None
            
    def deny(self):
        self.denied += 1


class Module():
    def __init__(self):
        self.checkouts = 0
        self.denied_count = 0
        self.max_checkouts = 0
        self.max_checkout_time = None
        self.users = defaultdict(User)
        self.pcs = defaultdict(int)
    
    def check_out(self, user, pc, timestamp):
        self.checkouts += 1
        if self.checkouts > self.max_checkouts:
            self.max_checkouts = self.checkouts
            self.max_checkout_time = timestamp
        self.users[user].check_out(timestamp)
        self.pcs[pc] += 1
    
    def check_in(self, user, pc, timestamp):
        self.users[user].check_in(timestamp)
        self.checkouts = max(0, self.checkouts - 1)
        
    def update(self, in_out, user, pc, timestamp):
        if in_out in ('OUT', 'CHECKOUT'):
            self.check_out(user, pc, timestamp)
        elif in_out in ('IN', 'CHECKIN'):
            self.check_in(user, pc, timestamp)
        elif in_out == 'DENIED':
            self.denied_count += 1
            self.users[user].deny()
        
    def reset(self, timestamp):
        self.checkouts = 0
        for user in self.users.values():
            user.check_in(timestamp)
        
    def __repr__(self):
        return "<Module, current checkouts %d, max %d at %s>" % (self.checkouts, self.max_checkouts, str(self.max_checkout_time))


def email_to_id(email: str) -> str:
    """Given an email address, resolve to a user id."""
    name_resolver.Set(ADSNameType.unknown, email)
    ldap_query = f'LDAP://{name_resolver.Get(ADSNameType.type_1779)}'
    ldap = win32com.client.GetObject(ldap_query)
    return ldap.get('CN')


def read_licenses():
    # line format is the same for both CST and MATLAB
    # 11:00:10 (cstd) OUT: "start" itx75623@DDAST0137
    line_regexp = re.compile(r'(?P<timestr>[ \d]\d:\d\d:\d\d) \(.*\) (?P<message>.*)')
    checkout_regexp = re.compile(r'(?P<in_out>IN|OUT|DENIED): "(?P<toolbox>.*?)" (?P<user>.*)?@(?P<pc>.*)  (?P<comment>)')
    if 'matlab' in sys.argv:
        log_path = r'\\apsv2\bjs54'
        log_name = 'FLEX.log'
        program_name = 'MATLAB'
        reset_message = 'License verification completed successfully.'
    else:  # Opera
        log_path = os.path.join(docs_folder, 'Opera Models')
        log_name = 'cst-lmgrd.log'  # get it from /tmp on apclara1
        program_name = 'Opera'
        reset_message = 'Starting vendor daemons ... '

    log_filename = os.path.join(log_path, log_name)
    print(log_filename)

    tool_usage = defaultdict(Module)
    user_usage = {}
    pc_usage = {}

    print(f"{program_name} license log analysis")
    start_date = datetime.datetime.now(datetime.UTC) - datetime.timedelta(days=90)
    print(f"Start date: {str(start_date)}")
    timestamp = None
    timezone = None

    for line in open(log_filename):
        if matches := line_regexp.match(line):
            timestr, message = matches.group('timestr', 'message')
            if message.startswith('TIMESTAMP'):  # m/d/y follows
                datestr = message[9:]
                timestamp = dateutil.parser.parse(datestr + ' ' + timestr)
                if timezone:
                    timestamp = timestamp.replace(tzinfo=timezone)
            elif message.startswith("Server's System Date and Time: "):
                datestr = message[31:]  # e.g. Tue Jan 07 2025 17:30:14 GMT follows
                timestamp = dateutil.parser.parse(datestr)
                timezone = timestamp.tzinfo

            if timestamp is None or timestamp < start_date:
                # before the log period we want, or we don't have date info yet
                continue
            time = dateutil.parser.parse(timestr)
            timestamp = timestamp.replace(hour=time.hour, minute=time.minute, second=time.second)  # combine with known date info

            if message == reset_message:
                # reset all licenses
                for tb in tool_usage.values():
                    tb.reset(timestamp)

            if matches := checkout_regexp.match(message):
                in_out, toolbox, user, pc, comment = matches.group('in_out', 'toolbox', 'user', 'pc', 'comment')
                if comment == 'FAIL':  # license denial
                    in_out = 'DENIED'
                if toolbox[-1:] == '_':  # skip entries ending _ as they are duplicates
                    continue
                tb = tool_usage[toolbox]
                tb.update(in_out, user.lower(), pc.lower(), timestamp)

    for name, toolbox in tool_usage.items():
        print(f'\n{name} {toolbox}')
        # First find the total length of each use
        for username, user in toolbox.users.items():
            user.checkout_lengths = [in_time - out_time for out_time, in_time in user.checkouts]
            user.total_len = sum(user.checkout_lengths, datetime.timedelta())

        user_list = sorted(toolbox.users.keys(), key=lambda name: toolbox.users[name].total_len, reverse=True)
        for username in user_list:
            user = toolbox.users[username]
            try:
                name_resolver.Set(3, 'clrc\\' + (username[:-2] if username.endswith('03') else username))
                ldap_query = f'LDAP://{name_resolver.Get(1)}'
                ldap = win32com.client.GetObject(ldap_query)
                email = ldap.get('mail')
            except:
                email = 'unknown'
            try:
                max_len = str(max(user.checkout_lengths))
                last_use = str(max([in_time for out_time, in_time in user.checkouts]))
            except ValueError:
                max_len = 'N/A'
                last_use = 'never'
            line = f'  {username} ({email}): {len(user.checkouts)} uses, total {str(user.total_len)}, {last_use=:}'
            if user.denied > 0:
                line += f', denied {user.denied} times'
            print(line)

        print('Usage by PC:', sorted(toolbox.pcs.items(), key=operator.itemgetter(1), reverse=True))
        print(f'Denied on {toolbox.denied_count:d} occasions')


if __name__ == '__main__':
    print(email_to_id('neil.thompson@stfc.ac.uk'))
