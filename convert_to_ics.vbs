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
objIcsFile, _ 
objTextFile, _ 
stderr, _ 
strDate, _ 
strEvent, _ 
strEventEndTime, _ 
strEventStartTime, _ 
strGardenEventTitle, _
strInputFilename, _ 
strLine, _ 
strOutputfile, _ 
strRecyclingEvent, _ 
strRecyclingEventTitle, _ 
strRefuseEventTitle, _ 
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

set objIcsFile = objFSO.OpenTextFile(strOutputfile, FOR_WRITING, true)

printHeader()

Do Until objTextFile.AtEndOfStream
	strLine = objTextFile.ReadLine
' Filter in the "data" lines
	If InStr(strLine, "-") > 0 Then
		arrData = lineToArray(strLine)
		strDate = arrData(0)
		strEvent = arrData(1)
' Write the iCal event for the particular date
		objIcsFile.writeline "BEGIN:VEVENT"
		objIcsFile.writeline "SUMMARY:" & rawToTitle(strEvent)
		objIcsFile.writeline "DTSTART;TZID=Europe/London:" & strDate & "T" & strEventStartTime & "00"
		objIcsFile.writeline "DTEND;TZID=Europe/London:" & strDate & "T" & strEventEndTime & "00"
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

Function lineToArray(line)
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

	lineToArray=array(dateValidDateMinusOneInIsoFormat,strRecyclingEvent)
End Function
