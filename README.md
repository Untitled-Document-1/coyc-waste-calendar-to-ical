# Waste Collection Calendar - plain text to iCal

This script was written to convert the City of York Council's Waste Collection Calendar (plain text version) into a format suitable for importing into online calendars such as [Google Calendar](https://calendar.google.com/).
The objective here is to get notified by email on the day or night before the bins/recyling is collected, as a reminder to _put the bin out_.
Most if not all online calendars provide the ability to send email notifications for upcoming events.

The `.ics` output file has been tested with Google Calendar and Yahoo Calendar.
Note: the script is a VBScript (Microsoft Windows only).

## How do I use this?

1. Go to the Waste Collection Calendar - https://myaccount.york.gov.uk/bin-collections - do a postcode lookup
2. Click the 'View your waste and recycling calendar for the current year' link
3. Click the 'TEXT ONLY VERSION' button
4. Copy all of the dates into the buffer, by left-click-dragging down the page and pressing CTRL+C
5. Paste into a plain text file with CTRL+V and save this to disk. This becomes your plain text input file
6. Download the `convert_to_ics.vbs` script somewhere
7. Open a Windows command prompt, and navigate to where you downloaded it: `chdir /d C:\path\to\file`
8. Type `cscript //nologo convert_to_ics.vbs /inputfile:"C:\path\to\file\sample_input_file.txt" /eventstarttime:"18:00" /eventendtime:"18:10" /reminderemailaddress:"me@example.org" /outputfile:"C:\path\to\file\ical_output_file.ics"`.
9. Adjust the named parameter values before pressing return on your keyboard.
10. This will generate an `.ics` file you can import into your calendar. Note: it's assumed that you want the event put in the day before the collection, as the collections are usually very early in the morning

A sample plain text input file is included for testing & comparison purposes.

## Tips

* Add a dedicated calendar for this, so that you can import the `.ics` file into this and set event notifications at the calendar-level. If, generally, you don't tend to use calendar event notifications, but would like to do so solely for bins/recyling events, then set the `reminderemailaddress` named parameter as in the example above. This will add a `VALARM` section to each event, so that you will be reminded by email. Otherwise you'll probably have to set notifications at the event-level, which would be a tedious, manual process.
* You can customise the calendar event titles using `/gardentitle`, `/refusetitle` & `/recyclingtitle` named parameters, for example, `/gardentitle:"Put the green bin out"`. Otherwise, the titles will be as they are in the input file; `GARDEN`,`REFUSE` & `RECYCLING`.
* The `/eventstarttime` & `/eventendtime` named parameters are for setting when the _put the bin out_ calendar event starts and ends. An event lasting 10 minutes would be a typical duration to use, as in the example `cscript ...` line above.

## Why not use recurring calendar events?

You could create three recurring events, one for `RECYCLING`, one for `REFUSE` and one for `GARDEN`, and it would likely accurately reflect the collection rota for your postcode. The only downside to this is that, if there are any anomolies around Christmas, where the collection day varies, the recurring event wouldn't reflect this. This script takes it from _the horses mouth_, so to speak, rather than assuming the pattern of weekly collections will always hold true.

## Disclaimer
* I am not employed or affiliated in any way with City of York Council
* I am not responsible if you ignore your emails and don't _put the bin out_ ;-)
