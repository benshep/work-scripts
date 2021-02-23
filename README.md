# Work scripts
Scripts that I use for work.
## join_zoom_meeting.pyw
Assign this to a hotkey (I use Win+Ctrl+Z) to join a Zoom meeting instantly, without having to find the appointment, click on the link, and have a tab open. Works whenever there is an Outlook appointment with a Zoom link in the subject or the body. Looks for the meeting that starts nearest to the current time. Obviously if you have two that start at the same time with Zoom links, it will make an arbitrary choice.

Command line arguments:

- `second` - if there are two meetings that start now, join the second one (order is arbitrary)
- `notes` - start a Markdown file with notes for the current meeting instead of joining it

## autohotkey.ahk
Hotkeys to activate or start various programs using [AutoHotKey](https://autohotkey.com/) on Windows, as well as some convenience functions.

