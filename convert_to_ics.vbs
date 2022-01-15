Option Explicit
Dim arrGbDate,arrGbDateWithoutDay,arrGbDateJustDate,arrGbDateJustNum,arrNights,colNamedArguments,objFSO,objTextFile,strInputFilename,strEventStartTime,strEventEndTime,strLine,strRecyclingEvent,dateValidDate,dateValidDateMinusOne,dateValidDateMinusOneInIsoFormat,stderr,strValidDate
Set colNamedArguments = WScript.Arguments.Named
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set stderr = objFSO.GetStandardStream(2)

If IsEmpty(colNamedArguments.Item("inputfile")) Then
    stderr.WriteLine "The full filepath to an input file must be specified"
    WScript.Quit 1
Else 
    strInputFilename = colNamedArguments.Item("inputfile")
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
    Const ForReading = 1
    Set objTextFile = objFSO.OpenTextFile(strInputFilename, ForReading)
Else
    stderr.WriteLine "The input file specified does not exist"
    WScript.Quit 1
End If

Wscript.Echo "BEGIN:VCALENDAR"
Wscript.Echo "PRODID://Shampoo//Calendar//EN"
Wscript.Echo "VERSION:2.0"

Do Until objTextFile.AtEndOfStream
    strLine = objTextFile.ReadLine
    If InStr(strLine, "-") > 0 Then
        
        'Example line:
        ' Wednesday, 8th December 2021 - RECYCLING
        arrNights = Split(strLine, " - ")
        
        ' - RECYCLING
        strRecyclingEvent = arrNights(1)
        
        ' Wednesday, 8th December 2021
        arrGbDate = Split(arrNights(0), ",")
        
        ' 8th December 2021
        arrGbDateWithoutDay = trim(arrGbDate(1))
        
        ' 8th
        arrGbDateJustDate = Split(arrGbDateWithoutDay, " ")
        
        ' 8
        arrGbDateJustNum = replace(replace(replace(replace(arrGbDateJustDate(0),"nd","",1,1),"st","",1,1),"rd","",1,1),"th","",1,1)
        
        ' 8 December 2021
        strValidDate = arrGbDateJustNum & " " & arrGbDateJustDate(1) & " " & arrGbDateJustDate(2)       
        
        dateValidDate = CDate(strValidDate)
        dateValidDateMinusOne = DateAdd("d",-1,dateValidDate)
        dateValidDateMinusOneInIsoFormat = Year(dateValidDateMinusOne) & Right("00" & Month(dateValidDateMinusOne), 2) & Right("00" & Day(dateValidDateMinusOne), 2)
        'Wscript.Echo DEBUG: dateValidDateMinusOneInIsoFormat & " " & strRecyclingEvent
        
        ' Write the iCal event for the particular date
        Wscript.Echo "BEGIN:VEVENT"
        Wscript.Echo "SUMMARY:" & strRecyclingEvent
        Wscript.Echo "DTSTART;TZID=Europe/London:" & dateValidDateMinusOneInIsoFormat & "T" & strEventStartTime & "00"
        Wscript.Echo "DTEND;TZID=Europe/London:" & dateValidDateMinusOneInIsoFormat & "T" & strEventEndTime & "00"
        Wscript.Echo "END:VEVENT"
    End If
Loop
objTextFile.Close
Wscript.Echo "END:VCALENDAR"