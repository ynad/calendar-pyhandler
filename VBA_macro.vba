'
' Sample VBA macro definitions to be used to integrate Excel spreadsheets.
'
' https://github.com/ynad/caldav-py-handler
' info@danielevercelli.it
'


' macro to add events of training/formation. Single or multi-day. Single or multiple attendants. Hours might be provided too
Sub calAdd_training()

    ' compose event name
    If IsEmpty(Cells(5, ActiveCell.Column - 1)) Then
        ' date and deadline same column & cell
        EventName = (ActiveSheet.Name + " " + Cells(5, ActiveCell.Column))
    Else
        ' column is splitted in date and deadline
        EventName = (ActiveSheet.Name + " " + Cells(5, ActiveCell.Column - 1))
    End If
    EventName = Application.WorksheetFunction.Trim(EventName)
    
    ' date format can be: single date "dd/mm/YYYY" or range of dates "dd/mm/YYYY-dd/mm/YYYY" or comma separated dates "dd/mm/YYYY,dd/mm/YYYY,..." (no white spaces)
    EventDate = ActiveCell.Value
    date_array = Split(EventDate, "-", 2)
    ' in case of 2 dates, multi-day event
    If UBound(date_array) > 0 Then
        start_day = date_array(0)
        end_day = date_array(1)
    ' in case of single date, single day event
    Else
        start_day = date_array(0)
        end_day = date_array(0)
    End If
    
    ' search for same event date, and add all corresponding values to string variable
    Dim N As Long, i As Long, j As Long
    Dim values_string As String
    ' count number of rows of current A column in current sheet
    N = Cells(Rows.Count, "A").End(xlUp).Row
    ' start position i for cells to start validating data on. 7 is related to the current example rows
    For i = 7 To N
        ' check if this cycle cell has the same value of the selected cell
        If Cells(i, ActiveCell.Column).Value = ActiveCell.Value Then
            ' check if data is valid, in this example if both ID are under "1000"
            If Cells(i, 2).Value < 1000 Then
                If Cells(ActiveCell.Row, 2) < 1000 Then
                    ' add data to string and append a comma
                    values_string = values_string & Cells(i, 3).Value & ", "
                End If
            ' otherwise, in this example, if both ID are over "1000"
            Else
                If Cells(ActiveCell.Row, 2) > 999 Then
                    ' add data to string and append a comma
                    values_string = values_string & Cells(i, 3).Value & ", "
                End If
            End If
        End If
    Next i
    
    ' compose event description
    ' use different pattern depending on the same parameter used before, in this example person ID number
    If Cells(ActiveCell.Row, 2) < 1000 Then
        EventDescr = Cells(5, ActiveCell.Column - 1) + "; PATTERN1: " + values_string
    Else
        EventDescr = Cells(5, ActiveCell.Column - 1) + "; PATTERN2: " + values_string
    End If
    ' trim extra white spaces in description
    EventDescr = Application.WorksheetFunction.Trim(EventDescr)

    ' DISABLED
    ' determine if course is marked to be taken - performs sub-string search. Key word is "TO DO"
    'i = InStr(ActiveCell.Value, "TO DO")
    ' if key word is found, grep preceding string(s) as event date(s)
    'If i <> 0 Then
    '    EventDate = Mid(ActiveCell.Value, 1, i - 2)
    'Else
    '    EventDate = ActiveCell.Value
    'End If

    ' call external script passing parameters in proper syntax
    Call Shell("python caldav-ics-client.py" & " --name " & """" & EventName & """" & " --descr " & """" & EventDescr & """" & " --start_day " & start_day & " --end_day " & end_day & " --cal " & """" & "mycalendar" & """" & " --alarm_type DISPLAY" & " --alarm_format D" & " --alarm_time 60" & """")

End Sub


' macro to add events of maintenance routines
Sub calAdd_maintenance()

    ' compose event name
    EventName = (ActiveSheet.Name + " " + Cells(ActiveCell.Row, 1))

    ' compose event description
    ' date and description same row
    If Not IsEmpty(Cells(ActiveCell.Row, 1)) Then
        EventDescr = Cells(ActiveCell.Row, 1)
    ' seatch in row -1
    ElseIf Not IsEmpty(Cells(ActiveCell.Row - 1, 1)) Then
        EventDescr = Cells(ActiveCell.Row - 1, 1)
    ' seatch in row -2
    ElseIf Not IsEmpty(Cells(ActiveCell.Row - 2, 1)) Then
        EventDescr = Cells(ActiveCell.Row - 2, 1)
    End If

    ' trim extra white spaces
    EventName = Application.WorksheetFunction.Trim(EventName)
    EventDescr = Application.WorksheetFunction.Trim(EventDescr)
    
    ' start and end day coincide. Hours might be provided too
    start_day = ActiveCell.Value
    end_day = ActiveCell.Value
    
    ' call external script passing parameters in proper syntax
    Call Shell("python caldav-ics-client.py" & " --name " & """" & EventName & """" & " --descr " & """" & EventDescr & """" & " --start_day " & start_day & " --end_day " & end_day & " --cal " & """" & "mycalendar" & """" & " --alarm_type DISPLAY" & " --alarm_format D" & " --alarm_time 1" & """")
    
End Sub


' macro to add events of maintenance routines
Sub calAdd_maintenance_invite()

    ' compose event name
    EventName = (ActiveSheet.Name + " " + Cells(ActiveCell.Row, 1))

    ' compose event description
    ' date and description same row
    If Not IsEmpty(Cells(ActiveCell.Row, 1)) Then
        EventDescr = Cells(ActiveCell.Row, 1)
    ' seatch in row -1
    ElseIf Not IsEmpty(Cells(ActiveCell.Row - 1, 1)) Then
        EventDescr = Cells(ActiveCell.Row - 1, 1)
    ' seatch in row -2
    ElseIf Not IsEmpty(Cells(ActiveCell.Row - 2, 1)) Then
        EventDescr = Cells(ActiveCell.Row - 2, 1)
    End If

    ' trim extra white spaces
    EventName = Application.WorksheetFunction.Trim(EventName)
    EventDescr = Application.WorksheetFunction.Trim(EventDescr)
    
    ' start and end day coincide. Hours might be provided too
    start_day = ActiveCell.Value
    end_day = ActiveCell.Value
    
    ' invite email(s). Must be separated by spaces
    ' search for correct given key word to identify correct cell in given possibilities
    If StrComp(Cells(6, 1), "TO BE INVITED") = 0 Then
        InviteEmail = Cells(6, 4)
    ElseIf StrComp(Cells(7, 1), "TO BE INVITED") = 0 Then
        InviteEmail = Cells(7, 4)
    ElseIf StrComp(Cells(5, 1), "TO BE INVITED") = 0 Then
        InviteEmail = Cells(5, 4)
    End If
        
    ' call external script passing parameters in proper syntax
    Call Shell("python caldav-ics-client.py" & " --name " & """" & EventName & """" & " --descr " & """" & EventDescr & """" & " --start_day " & start_day & " --end_day " & end_day & " --cal " & """" & "mycalendar" & """" & " --invite " & """" & InviteEmail & """" & " --alarm_type DISPLAY" & " --alarm_format D" & " --alarm_time 1" & """")
    
End Sub


' macro to iterate over a predefined set of cells and copy/fill values on given conditions
Sub Deadline()
    For Each c In Worksheets("One").Range("B10:B12").Cells
        If c.Value > 0 Then Worksheets("SUMMARY").Range("D5").Value = c.Value
    Next
    For Each c In Worksheets("Two").Range("C10:C12").Cells
        If c.Value > 0 Then Worksheets("SUMMARY").Range("C5").Value = c.Value
    Next
    For Each c In Worksheets("Three").Range("D10:D12").Cells
        If c.Value > 0 Then Worksheets("SUMMARY").Range("B5").Value = c.Value
    Next
End Sub


'
Sub Range()
    Dim rng As Range: Set rng = Application.Range("Sheet1!B10:C12")
    Dim cel As Range
    For Each cel In rng.Cells
        With cel
            Debug.Print .Address & ":" & .Value
        End With
    Next cel
End Sub


' macro trigger with double click on any cell inside given range
Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Application.Intersect(Target, Range("A1:E21")) Is Nothing Then
        EventName = (ActiveSheet.Name + " " + Range("A10").Value)
        Call Shell("python caldav-ics-client.py" & " --name " & EventName & " --start_day " & Target.Value & " --end_day " & Target.Value)
    End If
End Sub
