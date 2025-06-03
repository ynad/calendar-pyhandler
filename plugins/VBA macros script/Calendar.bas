Attribute VB_Name = "Calendar"
'
' VBA macro definitions to be used to integrate Excel spreadsheets.
'
''' Implementation example '''
'
' v0.7.1 - 2025.05.28
' https://github.com/ynad/caldav-py-handler
' info@danielevercelli.it
'


' macro to add events of training/formation. Single or multi-day. Hours might be provided too
Sub calAdd_training()

    scriptFolder = "C:\MY_SCRIPTS\calendar-pyhandler\"
    Debug.Print "Script folder: " & scriptFolder
    
    ' check if folder script exists on local PC
    If Not Dir(scriptFolder, vbDirectory) <> "" Then
        MsgBox "Script non trovato su questo PC, non posso procedere", vbCritical
        Exit Sub
    End If

    ' compose event name
    If Not IsEmpty(Cells(5, ActiveCell.Column)) Then
        ' date and deadline same column & cell
        EventName = ActiveSheet.Name + " " + Cells(5, ActiveCell.Column) + " " + Cells(6, ActiveCell.Column)
    Else
        ' column is splitted in date and deadline
        EventName = ActiveSheet.Name + " " + Cells(5, ActiveCell.Column - 1) + " " + Cells(6, ActiveCell.Column)
    End If
    EventName = Application.WorksheetFunction.Trim(EventName)
    
    ' DATE FORMAT can be:
    ' single date "dd/mm/YYYY"
    ' range of dates "dd/mm/YYYY - dd/mm/YYYY"
    ' list of dates, single or ranges "dd/mm/YYYY; dd/mm/YYYY - dd/mm/YYYY"
    '
    ' HOURS format can be:
    ' single date with fixed hours "dd/mm/YYYY, hh:mm _ hh:mm"
    ' range of dates, with first day start hour and last day end hour "dd/mm/YYYY - dd/mm/YYYY, hh:mm _ hh:mm"
    ' list of dates, with an hours duration each "dd/mm/YYYY, hh:mm _ hh:mm; dd/mm/YYYY - dd/mm/YYYY, hh:mm _ hh:mm"
    '
    ' the mandatory separator format are as follow, whitespaces included:
    ' " - " for range of dates
    ' "; " for lists
    ' ", " between day and hour
    ' " _ " between start and end hours

    ' set current selected cell as event date, then analyze it
    EventDate = ActiveCell.Value
    ' check if list of dates
    date_array = Split(EventDate, "; ", -1)
    If UBound(date_array) > 0 Then
    
        ' cylce over list and check if single or range dates
        Dim j As Long
        For j = 0 To UBound(date_array)

            ' search last string, when j=(len of array), for comment between "()", if found consider only date
            If j = UBound(date_array) Then
                date_comment = Split(date_array(j), " (", -1)
                If UBound(date_comment) > 0 Then
                    date_j = date_comment(0)
                Else
                    date_j = date_array(j)
                End If
            Else
                date_j = date_array(j)
            End If

            date_hours = Split(date_j, ", ", 2)
            ' date with fixed hours
            If UBound(date_hours) > 0 Then
                ' search correct format
                hours = Split(date_hours(1), " _ ", 2)
                If UBound(hours) > 0 Then
                    ' set start and end hours
                    start_hr = start_hr & " " & hours(0)
                    end_hr = end_hr & " " & hours(1)
                    ' re-set data to remaining string without hours
                    date_j = date_hours(0)
                End If
            Else
                ' set empty start and end hours
                start_hr = start_hr & " " & "00:00"
                end_hr = end_hr & " " & "00:00"
            End If
        
            date_range = Split(date_j, " - ", 2)
            ' date range
            If UBound(date_range) > 0 Then
                start_day = start_day & " " & date_range(0)
                end_day = end_day & " " & date_range(1)
            Else
            ' single date
                start_day = start_day & " " & date_range(0)
                end_day = end_day & " " & date_range(0)
            End If
        Next j

    Else
    ' else data range or single date

        ' check if hours were provided
        date_hours = Split(EventDate, ", ", 2)
        ' date with fixed hours
        If UBound(date_hours) > 0 Then
            ' search correct format
            hours = Split(date_hours(1), " _ ", 2)
            If UBound(hours) > 0 Then
                ' set start and end hours
                start_hr = hours(0)
                end_hr = hours(1)
                ' re-set data to remaining string without hours
                EventDate = date_hours(0)
            End If
        End If

        ' else check if date range
        date_array = Split(EventDate, " - ", 2)
        If UBound(date_array) > 0 Then
            start_day = date_array(0)
            ' search string for comment between "()", if found consider only date
            date_comment = Split(date_array(1), " (", -1)
            If UBound(date_comment) > 0 Then
                end_day = date_comment(0)
            Else
                end_day = date_array(1)
            End If
        Else
        ' single date
            date_array = Split(EventDate, " (", -1)
            If UBound(date_array) > 0 Then
                start_day = date_array(0)
                end_day = date_array(0)
            Else
                start_day = EventDate
                end_day = EventDate
            End If
        End If
    End If
    
    ' search for same event date, and add all corresponding values to string variable
    Dim N As Long, i As Long
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
    ' crop last comma+whitespace from values string
    values_string = Left(values_string, Len(values_string) - 2)
    
    ' compose event description
    ' use different pattern depending on the same parameter used before, in this example person ID number
    If Cells(ActiveCell.Row, 2) < 1000 Then
        If Not IsEmpty(Cells(5, ActiveCell.Column)) Then
            EventDescr = Cells(6, ActiveCell.Column) + " " + Cells(5, ActiveCell.Column) + vbCrLf + vbCrLf + "CAT1: " + vbCrLf + values_string
        Else
            EventDescr = Cells(6, ActiveCell.Column) + " " + Cells(5, ActiveCell.Column - 1) + vbCrLf + vbCrLf + "CAT1: " + vbCrLf + values_string
        End If
    Else
        If Not IsEmpty(Cells(5, ActiveCell.Column)) Then
            EventDescr = Cells(6, ActiveCell.Column) + " " + Cells(5, ActiveCell.Column) + vbCrLf + vbCrLf + "CAT2: " + vbCrLf + values_string
        Else
            EventDescr = Cells(6, ActiveCell.Column) + " " + Cells(5, ActiveCell.Column - 1) + vbCrLf + vbCrLf + "CAT2: " + vbCrLf + values_string
        End If
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
    'Create a new Shell Object
    Set objShell = VBA.CreateObject("Wscript.shell")
         
    'Provide the file path to the Python Exe
    PythonExe = """python.exe"""
         
    'Provide the file path to the Python script
    PythonScript = scriptFolder & "calendar-pyCLIent.py"
    
    ' add commands
    shellCommand = " --config " & scriptFolder & "user_settings.json" & " --name " & """" & EventName & """" & " --descr " & """" & EventDescr & """" & " --start_day " & """" & start_day & """" & " --end_day " & """" & end_day & """" & " --start_hr " & """" & start_hr & """" & " --end_hr " & """" & end_hr & """" & " --cal " & """" & "personal" & """" & " --alarm_type DISPLAY" & " --alarm_format D" & " --alarm_time 60" & """"
   
    ' run it hidden (0) and wait result (True)
    ' https://www.vbsedit.com/html/6f28899c-d653-4555-8a59-49640b0e32ea.asp
    exitCode = objShell.Run(PythonExe & PythonScript & shellCommand, 0, True)

    If exitCode <> 0 Then
        Debug.Print "Codice di ritorno d'errore, non continuo: " & exitCode
        MsgBox "Errore in esecuzione script: " & exitCode, vbCritical
    Else
        Debug.Print "Completato calAdd_training"
    End If

End Sub


' macro to add events of maintenance routines
Sub calAdd_maintenance()

    scriptFolder = "C:\MY_SCRIPTS\calendar-pyhandler\"
    Debug.Print "Script folder: " & scriptFolder
    
    ' check if folder script exists on local PC
    If Not Dir(scriptFolder, vbDirectory) <> "" Then
        MsgBox "Script non trovato su questo PC, non posso procedere", vbCritical
        Exit Sub
    End If

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
    
    ' set current selected cell as event date, then analyze it
    start_day = ActiveCell.Value
    end_day = ActiveCell.Value
    
    ' check if hours were provided
    date_hours = Split(start_day, ", ", 2)
    ' date with fixed hours
    If UBound(date_hours) > 0 Then
        ' search correct format
        hours = Split(date_hours(1), " _ ", 2)
        If UBound(hours) > 0 Then
            ' set start and end hours
            start_hr = hours(0)
            end_hr = hours(1)
            ' re-set data to remaining string without hours
            start_day = date_hours(0)
            end_day = date_hours(0)
        End If
    End If

    ' call external script passing parameters in proper syntax
    'Create a new Shell Object
    Set objShell = VBA.CreateObject("Wscript.shell")
         
    'Provide the file path to the Python Exe
    PythonExe = """python.exe"""
         
    'Provide the file path to the Python script
    PythonScript = scriptFolder & "calendar-pyCLIent.py"
    
    ' add commands
    shellCommand = " --config " & scriptFolder & "user_settings.json" & " --name " & """" & EventName & """" & " --descr " & """" & EventDescr & """" & " --start_day " & """" & start_day & """" & " --end_day " & """" & end_day & """" & " --start_hr " & """" & start_hr & """" & " --end_hr " & """" & end_hr & """" & " --cal " & """" & "personal" & """" & " --alarm_type DISPLAY" & " --alarm_format D" & " --alarm_time 1" & """"

    ' run it hidden (0) and wait result (True)
    ' https://www.vbsedit.com/html/6f28899c-d653-4555-8a59-49640b0e32ea.asp
    exitCode = objShell.Run(PythonExe & PythonScript & shellCommand, 0, True)

    If exitCode <> 0 Then
        Debug.Print "Codice di ritorno d'errore, non continuo: " & exitCode
        MsgBox "Errore in esecuzione script: " & exitCode, vbCritical
    Else
        Debug.Print "Completato calAdd_maintenance, exitCode: " & exitCode
    End If
    
End Sub


' macro to add events of maintenance routines
Sub calAdd_maintenance_invite()

    scriptFolder = "C:\MY_SCRIPTS\calendar-pyhandler\"
    Debug.Print "Script folder: " & scriptFolder
    
    ' check if folder script exists on local PC
    If Not Dir(scriptFolder, vbDirectory) <> "" Then
        MsgBox "Script non trovato su questo PC, non posso procedere", vbCritical
        Exit Sub
    End If

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
    
    ' set current selected cell as event date, then analyze it
    start_day = ActiveCell.Value
    end_day = ActiveCell.Value
    
    ' check if hours were provided
    date_hours = Split(start_day, ", ", 2)
    ' date with fixed hours
    If UBound(date_hours) > 0 Then
        ' search correct format
        hours = Split(date_hours(1), " _ ", 2)
        If UBound(hours) > 0 Then
            ' set start and end hours
            start_hr = hours(0)
            end_hr = hours(1)
            ' re-set data to remaining string without hours
            start_day = date_hours(0)
            end_day = date_hours(0)
        End If
    End If
    
    ' invite email(s). Must be separated by spaces
    ' search for correct given key word to identify correct cell in given possibilities
    If StrComp(Cells(6, 1), "TO BE INVITED") = 0 Then
        InviteEmail = Cells(6, 4)
    ElseIf StrComp(Cells(7, 1), "TO BE INVITED") = 0 Then
        InviteEmail = Cells(7, 4)
    ElseIf StrComp(Cells(5, 1), "TO BE INVITED") = 0 Then
        InviteEmail = Cells(5, 4)
    ElseIf StrComp(Cells(4, 1), "TO BE INVITED") = 0 Then
        InviteEmail = Cells(4, 4)
    End If
        
    ' call external script passing parameters in proper syntax
    'Create a new Shell Object
    Set objShell = VBA.CreateObject("Wscript.shell")
         
    'Provide the file path to the Python Exe
    PythonExe = """python.exe"""
         
    'Provide the file path to the Python script
    PythonScript = scriptFolder & "calendar-pyCLIent.py"
    
    ' add commands
    shellCommand = " --config " & scriptFolder & "user_settings.json" & " --name " & """" & EventName & """" & " --descr " & """" & EventDescr & """" & " --start_day " & """" & start_day & """" & " --end_day " & """" & end_day & """" & " --start_hr " & """" & start_hr & """" & " --end_hr " & """" & end_hr & """" & " --cal " & """" & "personal" & " --invite " & """" & InviteEmail & """" & """" & " --alarm_type DISPLAY" & " --alarm_format D" & " --alarm_time 1" & """"

    ' run it hidden (0) and wait result (True)
    ' https://www.vbsedit.com/html/6f28899c-d653-4555-8a59-49640b0e32ea.asp
    exitCode = objShell.Run(PythonExe & PythonScript & shellCommand, 0, True)

    If exitCode <> 0 Then
        Debug.Print "Codice di ritorno d'errore, non continuo: " & exitCode
        MsgBox "Errore in esecuzione script: " & exitCode, vbCritical
    Else
        Debug.Print "Completato calAdd_maintenance_invite"
    End If

End Sub

