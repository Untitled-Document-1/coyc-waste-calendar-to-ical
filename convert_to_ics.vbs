' Code formatted with http://www.vbindent.com
Option Explicit
Dim arrData, _
arrGbDate, _
arrGbDateJustDate, _
arrGbDateJustNum, _
arrGbDateWithoutDay, _
arrNights, _
colNamedArguments, _
dateValidDate, _
dateValidDateMinusOne, _
dateValidDateMinusOneInIsoFormat, _
objFSO, _
objRecycleInfo, _
objIcsFile, _
objTextFile, _
stderr, _
strAlarmTrigger, _
strDate, _
strEvent, _
strEventEndTime, _
strEventStartTime, _
strGardenEventTitle, _
strInputFilename, _
strLine, _
strOutputfile, _
strRecyclingEventTitle, _
strRefuseEventTitle, _
strReminderEmailAddress, _
strValidDate

Const FOR_READING = 1, FOR_WRITING = 2
Set colNamedArguments = WScript.Arguments.Named
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set stderr = objFSO.GetStandardStream(2)

If IsEmpty(colNamedArguments.Item("inputfile")) Then
    stderr.WriteLine "The full filepath to an input file must be specified"
    WScript.Quit 1
Else
    strInputFilename = colNamedArguments.Item("inputfile")
End If

If IsEmpty(colNamedArguments.Item("outputfile")) Then
    stderr.WriteLine "The full filepath to an output file must be specified"
    WScript.Quit 1
Else
    strOutputfile = colNamedArguments.Item("outputfile")
End If

If IsEmpty(colNamedArguments.Item("eventStartTime")) Then
    stderr.WriteLine "An event start time must be specified, e.g. 1800"
    WScript.Quit 1
Else
    strEventStartTime = colNamedArguments.Item("eventStartTime")
End If

If IsEmpty(colNamedArguments.Item("eventEndTime")) Then
    stderr.WriteLine "An event end time must be specified, e.g. 1810"
    WScript.Quit 1
Else
    strEventEndTime = colNamedArguments.Item("eventEndTime")
End If

If objFSO.FileExists(strInputFilename) Then
    Set objTextFile = objFSO.OpenTextFile(strInputFilename, FOR_READING)
Else
    stderr.WriteLine "The input file specified does not exist"
    WScript.Quit 1
End If

If IsEmpty(colNamedArguments.Item("refusetitle")) Then
    strRefuseEventTitle = "REFUSE"
Else
    strRefuseEventTitle = colNamedArguments.Item("refusetitle")
End If

If IsEmpty(colNamedArguments.Item("recyclingtitle")) Then
    strRecyclingEventTitle = "RECYCLING"
Else
    strRecyclingEventTitle = colNamedArguments.Item("recyclingtitle")
End If

If IsEmpty(colNamedArguments.Item("gardentitle")) Then
    strGardenEventTitle = "GARDEN"
Else
    strGardenEventTitle = colNamedArguments.Item("gardentitle")
End If

If colNamedArguments.Exists("reminderemailaddress") Then
    strReminderEmailAddress = colNamedArguments.Item("reminderemailaddress")
    If InStr(1, strReminderEmailAddress, "@yahoo.", VBTextCompare) > 0 Then
      strAlarmTrigger = "-PT5M"
    Else
      strAlarmTrigger = "-PT0M"
    End If
End If

set objIcsFile = objFSO.OpenTextFile(strOutputfile, FOR_WRITING, true)

printHeader()

Do Until objTextFile.AtEndOfStream
    strLine = objTextFile.ReadLine
    Set objRecycleInfo = New RecyclingEvent
    ' Write the iCal event for the particular date
    If objRecycleInfo.RecyclingEventType <> "" And objRecycleInfo.RecyclingEventDate <> "" Then
        objIcsFile.writeline "BEGIN:VEVENT"
        objIcsFile.writeline "SUMMARY:" & objRecycleInfo.RecyclingEventType
        objIcsFile.writeline "UID:Ical" & RandomString(32,objRecycleInfo.RecyclingEventDate & replace(strEventStartTime,":",""))
        objIcsFile.writeline "DTSTAMP;TZID=Europe/London:" & iso8601Date(Now)
        objIcsFile.writeline "DTSTART;TZID=Europe/London:" & objRecycleInfo.RecyclingEventDate & "T" & replace(strEventStartTime,":","") & "00"
        objIcsFile.writeline "DTEND;TZID=Europe/London:" & objRecycleInfo.RecyclingEventDate & "T" & replace(strEventEndTime,":","") & "00"
        If colNamedArguments.Exists("reminderemailaddress") Then
          objIcsFile.writeline "BEGIN:VALARM"
          objIcsFile.writeline "TRIGGER:" & strAlarmTrigger
          objIcsFile.writeline "ACTION:EMAIL"
          objIcsFile.writeline "ATTENDEE:" & strReminderEmailAddress
          objIcsFile.writeline "SUMMARY:Put the " & strRefuseEventTitle & " out"
          objIcsFile.writeline "DESCRIPTION:This is a reminder email about putting the " & strRefuseEventTitle & " out"
          objIcsFile.writeline "END:VALARM"
        End If
        objIcsFile.writeline "END:VEVENT"
    End If
Loop

printFooter()

objTextFile.Close
objIcsFile.Close

Sub printHeader()
    objIcsFile.writeline "BEGIN:VCALENDAR"
    objIcsFile.writeline "PRODID://Shampoo//Calendar//EN"
    objIcsFile.writeline "VERSION:2.0"
    objIcsFile.writeline "CALSCALE:GREGORIAN"
    objIcsFile.writeline "BEGIN:VTIMEZONE"
    objIcsFile.writeline "TZID:Europe/London"
    objIcsFile.writeline "BEGIN:DAYLIGHT"
    objIcsFile.writeline "TZOFFSETFROM:+0000"
    objIcsFile.writeline "TZOFFSETTO:+0100"
    objIcsFile.writeline "TZNAME:BST"
    objIcsFile.writeline "DTSTART:19700329T010000"
    objIcsFile.writeline "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU"
    objIcsFile.writeline "END:DAYLIGHT"
    objIcsFile.writeline "BEGIN:STANDARD"
    objIcsFile.writeline "TZOFFSETFROM:+0100"
    objIcsFile.writeline "TZOFFSETTO:+0000"
    objIcsFile.writeline "TZNAME:GMT"
    objIcsFile.writeline "DTSTART:19701025T020000"
    objIcsFile.writeline "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU"
    objIcsFile.writeline "END:STANDARD"
    objIcsFile.writeline "X-WR-TIMEZONE:Europe/London"
    objIcsFile.writeline "END:VTIMEZONE"
End Sub

Sub printFooter()
    objIcsFile.writeline "END:VCALENDAR"
End Sub

Function rawToTitle(raw)
    Select Case raw
    Case "REFUSE"
        rawToTitle = strRefuseEventTitle
    Case "RECYCLING"
        rawToTitle = strRecyclingEventTitle
    Case "GARDEN"
        rawToTitle = strGardenEventTitle
    End Select
End Function

Function lineToData(line,prop)
    'Example line:
    ' Wednesday, 8th December 2021 - RECYCLING
    Dim objDateMatches, _
    objEventMatches, _
    objMatch, _
    objRegExpDate, _
    objRegExpEvent, _
    strRecyclingEvent

    If prop = "RecyclingEventDate" Then
        Set objRegExpDate = New RegExp
        objRegExpDate.Pattern = ",\s([0-9]{1,2})(th|nd|st|rd)\s(January|February|March|April|May|June|July|August|September|October|November|December)\s([0-9]{4})"
        Set objDateMatches = objRegExpDate.Execute(line)
        If objDateMatches.Count > 0 Then
            Set objMatch = objDateMatches(0)
            dateValidDate = CDate(objMatch.SubMatches(0) & " " & objMatch.SubMatches(2) & " " & objMatch.SubMatches(3))
            dateValidDateMinusOne = DateAdd("d",-1,dateValidDate)
            dateValidDateMinusOneInIsoFormat = Year(dateValidDateMinusOne) & Right("00" & Month(dateValidDateMinusOne), 2) & Right("00" & Day(dateValidDateMinusOne), 2)
            lineToData=dateValidDateMinusOneInIsoFormat
        End If
    End If

    If prop = "RecyclingEventType" Then
        Set objRegExpEvent = New RegExp
        objRegExpEvent.Pattern = "(GARDEN|REFUSE|RECYCLING)$"
        Set objEventMatches = objRegExpEvent.Execute(line)
        If objEventMatches.Count > 0 Then
            Set objMatch = objEventMatches(0)
            strRecyclingEvent = objMatch.SubMatches(0)
            lineToData=strRecyclingEvent
        End If
    End If

End Function

Class RecyclingEvent
    Public Property Get RecyclingEventType
        RecyclingEventType = lineToData(strLine,"RecyclingEventType")
    End Property
    Public Property Get RecyclingEventDate
        RecyclingEventDate = lineToData(strLine,"RecyclingEventDate")
    End Property
End Class

' From https://stackoverflow.com/a/18448889/1754517
Function iso8601Date(dt)
  Dim s
  s = datepart("yyyy",dt)
  s = s & RIGHT("0" & datepart("m",dt),2)
  s = s & RIGHT("0" & datepart("d",dt),2)
  s = s & "T"
  s = s & RIGHT("0" & datepart("h",dt),2)
  s = s & RIGHT("0" & datepart("n",dt),2)
  s = s & RIGHT("0" & datepart("s",dt),2)
  iso8601Date = s
End Function

' Adapted from https://stackoverflow.com/a/30116847/1754517
Function RandomString(ByVal strLen, seed)
    Dim str, min, max, i
    Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789"
    min = 1
    max = Len(LETTERS)
    ' Randomize statement is without any args based on system time, so the following is a quick hack to ensure randomness
    Randomize seed
    For i = 1 to strLen
        str = str & Mid( LETTERS, Int((max-min+1)*Rnd+min), 1 )
    Next
    RandomString = str
End Function
