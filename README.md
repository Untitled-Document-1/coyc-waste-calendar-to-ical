# Waste Collection Calendar - plain text to iCal

This script was written to convert the City of York Council's Waste Collection Calendar (plain text version) into a format suitable for import into online calendars such as [Google Calendar](https://calendar.google.com/).
The objective here is to get notified by email on the day or night before the bins/recyling is collected, as a reminder to _put the bin out_. 
Most if not all online calendars provide the ability to send email notifications for upcoming events.

Currently the output file generated appears to be incompatible with Yahoo Calendar - attempting to do so results in a generic, unhelpful _Failed to import X calendar_ message.
The output file _is_ importable into Google Calendar, however.
Note: the script is a VBScript (Windows only).

1. Go to the Waste Collection Calendar - https://myaccount.york.gov.uk/bin-collections - do a postcode lookup
2. Click the "View your waste and recycling calendar for the current year" link
3. Click the 'TEXT ONLY VERSION' button
4. Copy all of the dates into the buffer, by left-click-dragging down the page and pressing CTRL+C
5. Paste into a plain text file with CTRL+V and save this to disk
6. Download this script somewhere
7. Open a Windows command prompt
8. Type `cscript //nologo convert_to_ics.vbs /inputfile:"C:\path\to\file\sample_input_file.txt" /eventstarttime:"18:00" /eventendtime:"18:10" > ical_output_file.ics`
9. This will generate an `.ics` file you can import into your calendar. Note it's assumed that you want the event put in the day before the collection

A sample plain text input file is included for testing & comparison purposes.

## Tip(s):

* Add a dedicated calendar for this, so that you can set event notifications at the calendar-level. If, generally, you don't tend to use calendar event notifications, but would like to do so solely for bins/recyling events, then a separate calendar is probably for you. Otherwise you'll probably have to set notifications at the event-level, which would be a tedious, manual process.

## Why not use recurring calendar events?

You could create three recurring events, one for `RECYCLING`, one for `REFUSE` and one for `GARDEN`, and it would likely accurately reflect the collection rota for your postcode. The only downside to this is that, if there are any anomolies around Christmas, where the collection day varies, the recurring event wouldn't reflect this. This script takes it from _the horses mouth_, so to speak, rather than assuming the pattern of weekly collections will always hold true.


## TODO: 
* add missing fields per https://icalendar.org/validator.html - maybe this will fix the import failures seen when trying to import to Yahoo Calendar?
