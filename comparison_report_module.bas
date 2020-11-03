Attribute VB_Name = "comparison_report_module"
Option Explicit
'Ed Branigan 21/10/2020

Sub fill_Cols()

    Dim tms_ws As Worksheet
    Dim ros_ws As Worksheet
    Dim result_ws As Worksheet
    Dim last_row As Integer
    Dim last_col As Integer
    
    Dim i As Integer
    Dim y As Integer
    Dim ros_row As Integer
    Dim search_string As String
    
    Set tms_ws = ThisWorkbook.Worksheets("tms_report")
    Set ros_ws = ThisWorkbook.Worksheets("staging_sheet")
    Set result_ws = ThisWorkbook.Worksheets("results")
    last_row = result_ws.Cells(result_ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To last_row
    
        search_string = result_ws.Cells(i, 2).Text 'dcam
        ros_row = find_cell(search_string) 'row no. of dcam found
        last_col = ros_ws.Cells(1, ros_ws.Columns.Count).End(xlToLeft).Column
            For y = 5 To last_col
                
                If result_ws.Cells(i, 5).Text = ros_ws.Cells(1, y).Text Then

                        If ros_row <> 0 Then
                        
                            result_ws.Cells(i, 7).Value = ros_ws.Cells(ros_row, y).Value
                            
                        End If
                End If
            
            Next y
    Next i
    
End Sub

Function find_cell(search_string As String) As Integer

    Dim rng As Range
    Dim cell As Range
    
    Set rng = ThisWorkbook.Worksheets("staging_sheet").Columns("B:B")
    Set cell = rng.Find(What:=search_string, _
    LookIn:=xlFormulas, LookAt:=xlWhole, MatchCase:=False)
    
    If cell Is Nothing Then
        
    Else
        find_cell = cell.Row
    End If
    
End Function

Sub clean_col(col As Integer)
    
    Dim result_ws As Worksheet
    Dim x As Integer
    Dim trimmed_text As String
    Dim rng As Range
    Dim last_row As Integer
    Set result_ws = ThisWorkbook.Worksheets("results")
    last_row = result_ws.Cells(result_ws.Rows.Count, "B").End(xlUp).Row
    
    For x = 2 To last_row
        Set rng = result_ws.Cells(x, col)
        trimmed_text = rng.Text
        trimmed_text = Replace(trimmed_text, " ", "")
        trimmed_text = Replace(trimmed_text, ":", "")
        
        Select Case rng
            Case "OFF-OFF"
                trimmed_text = "OFF"
            Case "ALM-ALM"
                trimmed_text = "ALM"
            Case "HOL-HOL"
                trimmed_text = "HOL"
            Case "AAB-AAB"
                trimmed_text = "AAB"
            Case "SCK-SCK"
                trimmed_text = "SCK"
            Case "UAB-UAB"
                trimmed_text = "UAB"
        End Select
        
        rng = trimmed_text
        
    Next x

End Sub

Sub clean_shifts()

    clean_col (6)
    clean_col (7)
    
End Sub

Sub find_DCAMS()
    
    Dim tms_ws As Worksheet
    Dim ros_ws As Worksheet
    
    Dim x, y As Integer
    Dim name, surname As String
    
    Set tms_ws = ThisWorkbook.Worksheets("tms_report")
    Set ros_ws = ThisWorkbook.Worksheets("staging_sheet")
     
            Next y
            
        End If
    
    Next x
    
End Sub

Sub match_cols()

    Dim ws As Worksheet
    Dim x As Integer
    Dim last_row As Integer
    Set ws = Sheets("results")
    last_row = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For x = 2 To last_row
        
        If ws.Cells(x, 6) = ws.Cells(x, 7) Then
            ws.Cells(x, 8) = "match"
        Else
        End If
        
        If ws.Cells(x, 6) = "OFF" And ws.Cells(x, 7) = "AAB" Then
            ws.Cells(x, 8) = "match"
        Else
        End If
        
        Select Case ws.Cells(x, 6).Text
            Case "OFF"
                If ws.Cells(x, 7) = "OFF" Or ws.Cells(x, 7) = "HOL" Then
                    ws.Cells(x, 8) = "match"
                End If
            Case Else
                
        End Select
        
        If ws.Cells(x, 6) <> "OFF" And ws.Cells(x, 7) = "ALM" Then
            ws.Cells(x, 8) = "match"
        Else

        End If
        
    Next x
    
End Sub

Sub make_roster_table()

    Application.ScreenUpdating = False
    ActiveWindow.Visible = True
    Dim input_date As String
    input_date = ThisWorkbook.Worksheets("Options").Range("C2").Text
    
    If IsDate(input_date) Then
        Sheets("staging_sheet").Cells.Clear
        create_headers (input_date)
        get_roster "Roster1", (input_date)
        get_roster "Roster2", (input_date)
        get_roster "Roster3", (input_date)
        get_roster "Roster4", (input_date)
    Else
        MsgBox ("Please enter a valid date")
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub get_roster(roster_name As String, week_date As String)

    Dim roster_address As String
    Dim this_ws As Worksheet
    Dim roster_wb As Workbook
    Dim i As Integer
    
    roster_address = "X:\Planning\Warehouse Planning\Rosters\Individual\" & roster_name & ".xls"
    Set this_ws = ThisWorkbook.Worksheets("staging_sheet")
    Set roster_wb = Workbooks.Open(roster_address, UpdateLinks:=False, ReadOnly:=True)
    
    For i = 1 To roster_wb.Worksheets.Count
    
        If regex_name_tester(roster_wb.Worksheets(i).name) Then
        
            get_roster_lines roster_wb.Worksheets(i), week_date
            
        End If
        
    Next i
    
    roster_wb.Saved = True
    roster_wb.Close
        
End Sub

Sub create_headers(week_date As Date)
    
    Dim ws As Worksheet
    Dim no_of_weeks As Integer
    Dim i As Integer
    Dim week_days As Integer
    Dim posi As Integer
    Set ws = ThisWorkbook.Worksheets("staging_sheet")
    no_of_weeks = ThisWorkbook.Worksheets("Options").Range("H2")
    week_days = 0
    posi = 5
    
    For i = 1 To no_of_weeks

        ws.Cells(1, 1) = "ID"
        ws.Cells(1, 2) = "DCAM"
        ws.Cells(1, 3) = "Name"
        ws.Cells(1, 4) = "Surname"
        ws.Cells(1, posi) = Format(week_date, "dd/mm/yy")
        posi = posi + 1
        ws.Cells(1, posi) = Format(week_date + 1 + week_days, "dd/mm/yy")
        posi = posi + 1
        ws.Cells(1, posi) = Format(week_date + 2 + week_days, "dd/mm/yy")
        posi = posi + 1
        ws.Cells(1, posi) = Format(week_date + 3 + week_days, "dd/mm/yy")
        posi = posi + 1
        ws.Cells(1, posi) = Format(week_date + 4 + week_days, "dd/mm/yy")
        posi = posi + 1
        ws.Cells(1, posi) = Format(week_date + 5 + week_days, "dd/mm/yy")
        posi = posi + 1
        ws.Cells(1, posi) = Format(week_date + 6 + week_days, "dd/mm/yy")
        
        posi = posi + 1
        week_days = week_days + 7
        
    Next i

End Sub

Private Function regex_name_tester(workSheetName As String)

    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
      .Pattern = "^[0-9 ]+$"
    End With
    regex_name_tester = RegEx.test(workSheetName)

End Function

Sub get_roster_lines(ws As Worksheet, week_date As String)

    Dim current_row As Integer
    Dim date_row As Integer
    Dim found_cell As Range
    Dim this_ws As Worksheet
    Dim col_no As Integer
    Dim no_of_weeks As Integer
    Dim dt As Date
    dt = CDate(week_date)
    Dim x As Integer
    
    no_of_weeks = ThisWorkbook.Worksheets("Options").Range("H2").Value
    Set this_ws = ThisWorkbook.Worksheets("staging_sheet")
    current_row = this_ws.Cells(Rows.Count, "B").End(xlUp).Row + 1
    col_no = 0
    
    For x = 1 To (no_of_weeks)
    
        dt = dt + col_no
        Debug.Print (dt)
        Set found_cell = ws.Range("A:A").Find(Format(dt, "dd-mmm-yy"), LookIn:=xlValues)
    
        date_row = found_cell.Row
        this_ws.Cells(current_row, 2) = ws.Range("B7") 'Dcam
        this_ws.Cells(current_row, 3) = ws.Range("B5") 'Name
        this_ws.Cells(current_row, 4) = ws.Range("A5") 'Surname
    
        this_ws.Cells(current_row, 5 + col_no) = _
        Format(ws.Cells(date_row, 3), "hh:mm") & " - " & Format(ws.Cells(date_row, 4), "hh:mm")
        this_ws.Cells(current_row, 6 + col_no) = _
        Format(ws.Cells(date_row, 8), "hh:mm") & " - " & Format(ws.Cells(date_row, 9), "hh:mm")
        this_ws.Cells(current_row, 7 + col_no) = _
        Format(ws.Cells(date_row, 13), "hh:mm") & " - " & Format(ws.Cells(date_row, 14), "hh:mm")
        this_ws.Cells(current_row, 8 + col_no) = _
        Format(ws.Cells(date_row, 18), "hh:mm") & " - " & Format(ws.Cells(date_row, 19), "hh:mm")
        this_ws.Cells(current_row, 9 + col_no) = _
        Format(ws.Cells(date_row, 23), "hh:mm") & " - " & Format(ws.Cells(date_row, 24), "hh:mm")
        this_ws.Cells(current_row, 10 + col_no) = _
        Format(ws.Cells(date_row, 28), "hh:mm") & " - " & Format(ws.Cells(date_row, 29), "hh:mm")
        this_ws.Cells(current_row, 11 + col_no) = _
        Format(ws.Cells(date_row, 33), "hh:mm") & " - " & Format(ws.Cells(date_row, 34), "hh:mm")
        
        col_no = col_no + 7
        
    Next x
    
End Sub

