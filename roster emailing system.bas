Attribute VB_Name = "Btn_GetRosters"
Option Explicit
'Eddie Branigan 08/04/2020

'updated 20/04/20
''added specific dcam roster button and sub
''created email lookup for rf rosters

'updated 30/04/20
''added sorting method sub routine for ordering worksheets numerically
''created new visuals for button macros for ease of use
''created dropdown lists for ease of use
''created a message field to attach a short message with roster being emailed

'updated 11/05/20
''added a roster size selection for 5, 13, 26 week selections
''added shading to every second row of rosters emailed for better visual fidelity
''added confirmation box for sending emails button

'updated 12/05/20
''added date and weeks of roster to title of email sent
''fixed bug with not selecting all roster lines for emailing

Sub Btn_GetRosters()
       
    Dim rosterWB As Workbook
    Dim ThisWorkBook As Workbook
    Dim xWb As Worksheet
    Dim i As Integer
    Dim rosterSelected As String
    Dim cell As Range
    Set cell = Worksheets("Email Handler").Range("F2")
    rosterSelected = cell.Text
    
    'Clear all worksheets except for email handler
    DeleteSheets
    
    Application.ScreenUpdating = False
    
    '***************************************************************************************************************
    'Set rosterWB = Workbooks.Open("C:\Users\potato\" & rosterSelected & ".xls", _
    UpdateLinks:=False, ReadOnly:=True)
    '***************************************************************************************************************
    
    ActiveWindow.Visible = True
    
    Dim wsCount As Integer
    wsCount = rosterWB.Worksheets.Count

    For i = 1 To wsCount
        If regExNameTester(rosterWB.Worksheets(i).name) Then
            If rosterWB.Worksheets(i).Range("B7") <> "" _
            And rosterWB.Worksheets(i).Range("A5") <> "" _
            And rosterWB.Worksheets(i).Range("B5") <> "" Then
            
                getRosterInfo rosterWB, rosterWB.Worksheets(i).name
            End If
            
        End If
    Next i
     
    rosterWB.Saved = True
    rosterWB.Close
    
    Application.ScreenUpdating = True
    
    getEmails
    
    SortWorksheetsTabs
    
End Sub

Sub getRosterInfo(rosterNum As Workbook, rosSheetName As String)
    
    Dim dcam As String
    Dim email As String
    Dim name As String
    Dim row As String
    Dim DateRange As Range
    Dim selectedDate As String
    Dim posi As Integer
    Dim sdDate As Date
    Dim dcamRow As Range
    Dim dateVar As String
    Dim numberOfWeeks As Integer
    Dim x As Integer
    Dim wNo As Integer
    Dim testNameRange As Range
    row = ""
    
    dcam = rosterNum.Worksheets(rosSheetName).Range("B7").Value
    email = ""
    name = rosterNum.Sheets(rosSheetName).Range("B5").Value _
    & " " & rosterNum.Sheets(rosSheetName).Range("A5").Value
    
    With ThisWorkBook
        .Sheets.Add(After:=.Sheets("email handler")).name = dcam
    End With
    
    setTabColour ThisWorkBook.Sheets(dcam), rosterNum.Sheets(rosSheetName)
    
    'Column names and formatting of dcam worksheets
    ThisWorkBook.Worksheets(dcam).Range("A1") = "Name:"
    ThisWorkBook.Worksheets(dcam).Range("A2") = "Dcam:"
    ThisWorkBook.Worksheets(dcam).Range("A3") = "Email:"
    ThisWorkBook.Worksheets(dcam).Range("A1:A3").Font.Bold = True
    ThisWorkBook.Worksheets(dcam).Range("B1") = name
    ThisWorkBook.Worksheets(dcam).Range("B2") = dcam
    ThisWorkBook.Worksheets(dcam).Range("B3") = email
    'sunday
    ThisWorkBook.Worksheets(dcam).Range("B4:C4").Merge
    ThisWorkBook.Worksheets(dcam).Range("B4") = "Sunday"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("B4"))
    'monday
    ThisWorkBook.Worksheets(dcam).Range("D4:E4").Merge
    ThisWorkBook.Worksheets(dcam).Range("D4") = "Monday"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("D4"))
    'tuesday
    ThisWorkBook.Worksheets(dcam).Range("F4:G4").Merge
    ThisWorkBook.Worksheets(dcam).Range("F4") = "Tuesday"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("F4"))
    'wednesday
    ThisWorkBook.Worksheets(dcam).Range("H4:I4").Merge
    ThisWorkBook.Worksheets(dcam).Range("H4") = "Wednesday"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("H4"))
    'thursday
    ThisWorkBook.Worksheets(dcam).Range("J4:K4").Merge
    ThisWorkBook.Worksheets(dcam).Range("J4") = "Thursday"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("J4"))
    'friday
    ThisWorkBook.Worksheets(dcam).Range("L4:M4").Merge
    ThisWorkBook.Worksheets(dcam).Range("L4") = "Friday"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("L4"))
    'saturday
    ThisWorkBook.Worksheets(dcam).Range("N4:O4").Merge
    ThisWorkBook.Worksheets(dcam).Range("N4") = "Saturday"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("N4"))
    'hours
    ThisWorkBook.Worksheets(dcam).Range("P4") = "Hours"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("P4"))
    'overtime
    ThisWorkBook.Worksheets(dcam).Range("Q4") = "Overtime"
    formatDaysOfWeek (ThisWorkBook.Worksheets(dcam).Range("Q4"))
    
    'set to used range, quicker
    Set DateRange = rosterNum.Worksheets(rosSheetName).Range("A9:A300")
    sdDate = ThisWorkBook.Worksheets("Email Handler").Range("F5")
    posi = 5
    
    numberOfWeeks = setNumberOfWeeks

        For x = 1 To numberOfWeeks
            selectedDate = ""
            selectedDate = sdDate
            
            For Each dcamRow In DateRange
                dateVar = dcamRow

                    If dateVar = selectedDate Then
                        row = dcamRow.row
                        fetchRosterLines rosterNum, rosSheetName, row, CStr(posi), dcam
                    End If
            Next dcamRow
                    
            sdDate = sdDate + 7
            
            posi = posi + 1
        Next x
    
End Sub

Sub formatDaysOfWeek(day As Range)

    day.Font.Bold = True
    day.HorizontalAlignment = xlCenter

End Sub

Function regExNameTester(workSheetName As String)

    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")

    With RegEx
      .Pattern = "^[0-9 ]+$"
    End With

    regExNameTester = RegEx.Test(workSheetName)

End Function

Sub DeleteSheets()

    Dim xWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.name <> "Email Handler" Then
            xWs.Delete
        End If
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub fetchRosterLines(rosterNum As Workbook, rosSheetName As String, _
    row As String, posi As String, dcam As String)
        
    Dim dcamRow As Range
    Dim dateVar As String
    Dim LastRow As String
    Dim DateRange As Range
    Dim selectedDate As String
    Dim borderRange As Range
          
        'date cell
        ThisWorkBook.Worksheets(dcam).Range("A" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("A" & row)
        ThisWorkBook.Worksheets(dcam).Range("A" & posi).NumberFormat = "dd-mmm-yy;@"
        
        'sun
        ThisWorkBook.Worksheets(dcam).Range("B" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("C" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("C" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("D" & row)
    
        'mon
        ThisWorkBook.Worksheets(dcam).Range("D" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("H" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("E" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("I" & row)
           
        'tue
        ThisWorkBook.Worksheets(dcam).Range("F" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("M" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("G" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("N" & row)
        
        'wed
        ThisWorkBook.Worksheets(dcam).Range("H" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("R" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("I" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("S" & row)
        
        'thu
        ThisWorkBook.Worksheets(dcam).Range("J" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("W" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("K" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("X" & row)
        
        'fri
        ThisWorkBook.Worksheets(dcam).Range("L" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("AB" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("M" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("AC" & row)
        
        'sat
        ThisWorkBook.Worksheets(dcam).Range("N" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("AG" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("O" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("AH" & row)
        
        'total and overtime
        ThisWorkBook.Worksheets(dcam).Range("P" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("AL" & row)
        
        ThisWorkBook.Worksheets(dcam).Range("Q" & posi) _
        = rosterNum.Sheets(rosSheetName).Range("AM" & row)
        
        'format times and centre
        ThisWorkBook.Worksheets(dcam).Range("B" & posi & ":" & "O" & posi) _
        .NumberFormat = "hh:mm;@"
        ThisWorkBook.Worksheets(dcam).Range("A" & posi & ":" & "Q" & posi) _
        .HorizontalAlignment = xlCenter
        
        'need to change border size
        LastRow = ThisWorkBook.Worksheets(dcam).Cells(1048576, 1).End(xlUp).row
        Set borderRange = ThisWorkBook.Worksheets(dcam).Range("A4:Q" & LastRow)

        With borderRange.Borders
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlThin
        End With
        
     Dim Counter As Integer
    
       'For every row in the current selection...
        For Counter = 1 To borderRange.Rows.Count
            'If the row is an odd number (within the selection)...
            If Counter Mod 2 = 1 Then
                'Set the pattern to xlGray16.
                borderRange.Rows(Counter).Interior.Pattern = xlGray16
            End If
        Next
        
End Sub

Sub setTabColour(ws As Worksheet, rosterSheet As Worksheet)

    ws.Tab.ColorIndex = rosterSheet.Tab.ColorIndex
    
End Sub

Sub Btn_getSpecificRosters()
    
    Dim rosterWB As Workbook
    Dim xWb As Worksheet
    Dim i As Integer
    Dim wsCount As Integer
    Dim rosterSelected As String
    Dim rosterCell As Range
    Dim dcamRange As Range
    Dim dcamCell As Range
    
    Set dcamRange = Range("L2:L30")
    Set rosterCell = Range("F2")
    
    rosterSelected = rosterCell.Text
    
    'Clear all worksheets except for email handler
    DeleteSheets
    
    Application.ScreenUpdating = False

    '***************************************************************************************************************
    'Set rosterWB = Workbooks.Open("C:\Users\potato\" & rosterSelected & ".xls", _
    UpdateLinks:=False, ReadOnly:=True)
    '***************************************************************************************************************
    
    ActiveWindow.Visible = True
    
    wsCount = rosterWB.Worksheets.Count
     
    'go through all worksheets in roster workbook
    For i = 1 To wsCount
        'check if they are worksheets with numbers only as their name
        If regExNameTester(rosterWB.Worksheets(i).name) Then
            'if the worksheet hasn't got blank areas in dcam and name cells
            If rosterWB.Worksheets(i).Range("B7") <> "" _
            And rosterWB.Worksheets(i).Range("A5") <> "" _
            And rosterWB.Worksheets(i).Range("B5") <> "" Then
                'for all cells filled in, in L range of email handler
                For Each dcamCell In dcamRange
                    
                    If dcamCell.Value = rosterWB.Worksheets(i).Range("B7").Value Then
                        getRosterInfo rosterWB, rosterWB.Worksheets(i).name
                    End If
                    
                Next dcamCell
                
            End If
            
        End If
        
    Next i
    
    rosterWB.Saved = True
    rosterWB.Close
    
    Application.ScreenUpdating = True
    
    getEmails
    
End Sub

Sub getEmails()

    Application.ScreenUpdating = False

    Dim rfWB As Workbook
    Dim emailWS As Worksheet
    Dim wsCount As Integer
    Dim LastRow As Integer
    Dim emailRow As Integer
    Dim wsIndex As Integer
    
    '****************************************************************************************************
    Set rfWB = Workbooks.Open("C:\Users\potato\RFspreadsheet.xlsx", _
    UpdateLinks:=False, ReadOnly:=True)
    '****************************************************************************************************
    
    Set emailWS = rfWB.Sheets("email")
    
    LastRow = emailWS.Cells(emailWS.Rows.Count, "A").End(xlUp).row
    wsCount = ThisWorkBook.Worksheets.Count
    
    For wsIndex = 1 To wsCount
        If ThisWorkBook.Sheets(wsIndex).name <> "Email Handler" Then
            For emailRow = 1 To LastRow
                If emailWS.Range("A" & emailRow).Text = ThisWorkBook.Sheets(wsIndex).name Then
                    ThisWorkBook.Sheets(wsIndex).Range("B3") = emailWS.Range("B" & emailRow).Text
                End If
            Next emailRow
        End If
    Next wsIndex
    
    rfWB.Saved = True
    rfWB.Close
    
    Application.ScreenUpdating = True
    
End Sub

Sub SortWorksheetsTabs()

    Application.ScreenUpdating = False
    Dim ShCount As Integer, i As Integer, j As Integer
    ShCount = Sheets.Count

    For i = 1 To ShCount - 1
        For j = i + 1 To ShCount
            If UCase(Sheets(j).name) < UCase(Sheets(i).name) Then
                Sheets(j).Move Before:=Sheets(i)
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
    ThisWorkBook.Worksheets("email handler").Move Before:=ActiveWorkbook.Sheets(1)
    
End Sub

Function setNumberOfWeeks() As Integer
    
    Dim weeksSelected As String
    weeksSelected = ThisWorkBook.Worksheets("email handler").Range("F8")
    
    Select Case weeksSelected

        Case "5 weeks"
          setNumberOfWeeks = 5
          
        Case "10 weeks"
            setNumberOfWeeks = 10
        
        Case "13 weeks"
          setNumberOfWeeks = 13
        
        Case "26 weeks"
          setNumberOfWeeks = 26

    End Select

End Function

Sub BtnMailList()

    Dim name As String
    Dim dcam As String
    Dim email As String
    Dim i As Integer
    Dim k As Integer
    
    k = 2

    For i = 2 To ActiveWorkbook.Worksheets.Count
        If Worksheets(i).name <> "Email handler" Then

            name = Worksheets(i).Range("B1")
            dcam = Worksheets(i).Range("B2")
            email = Worksheets(i).Range("B3")

            Worksheets("Email handler").Range("A" & k) = dcam
            Worksheets("Email handler").Range("B" & k) = name
            Worksheets("Email handler").Range("D" & k) = email

        End If
        
        k = k + 1
    Next i
    
End Sub

Sub Bt_FillDateCol()
    
    Dim row As String
    Dim rng As Range

    row = Cells(Rows.Count, 1).End(xlUp).row
    
    Range("C2:C" & row) = Range("F5").Text
    
End Sub

'Eddie Branigan 06/04/2020

Sub Btn_EmailRoster()

        Dim dcam As String, emailAddr As String, workWeek As String
        Dim name As String, roster As String, msg As String
        Dim i As Long
        Dim rng As Range
        
        LastRow = Cells(Rows.Count, "A").End(xlUp).row
        msg = ThisWorkBook.Worksheets("email handler").Range("F25").Text
        
        
        answer = MsgBox("Are you sure you want to send these email/s?" _
        , vbQuestion + vbYesNo + vbDefaultButton2, "Eddie  says:")
        

        If answer = vbYes Then
            'iterate throgh rows to be sent
            For i = 2 To LastRow
                Worksheets("Email Handler").Select
                dcam = Cells(i, 1)
                name = Cells(i, 2)
                workWeek = Cells(6, 5)
    
                emailAddr = Cells(i, 4)
                Set rng = getRange(dcam)
                If emailAddr <> "" Then
                'Send the mail
                    Send_Mail emailAddr, workWeek, name, roster, rng, msg
                    
                    'Sets the cells to email sent
                    Worksheets("Email Handler").Select
                    Cells(i, 5).Interior.ColorIndex = 43
                    Cells(i, 5).Value = "Email Sent"
                Else
                    Worksheets("Email Handler").Select
                    Cells(i, 5).Interior.ColorIndex = 3
                    Cells(i, 5).Value = "Failed"
                End If
                
            Next i
        Else
            MsgBox ("No mail was sent")
        End If
    
End Sub

Sub Send_Mail(emailAdd As String, workWeek As String, name As String, roster As String, rng As Range, msg As String)
                
    Dim firstMsg As String
    Dim selectedDate As String
    Dim selectedWeeks As String
    Dim secMsg As String
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    selectedDate = ThisWorkBook.Worksheets("email handler").Range("F5").Text
    selectedWeeks = ThisWorkBook.Worksheets("email handler").Range("F8").Text
    
    firstMsg = "<h1>Your Roster from week: " & selectedDate & " for " & selectedWeeks & "</h1>"
    secMsg = "<p>For furter information, please contact planning.</p>"
    
    With OutMail
        .To = emailAdd
        .Subject = "Your Roster"
        .HTMLBody = firstMsg & RangetoHTML(rng) & secMsg
        .Send
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub

Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    rng.Copy
    
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
    With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=TempFile, _
        Sheet:=TempWB.Sheets(1).name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                              "align=left x:publishsource=")
    TempWB.Close SaveChanges:=False
    Kill TempFile
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
    
End Function

Function getRange(dcam As String) As Range
    
    Dim LastRow As String
    Dim posi As Integer
    LastRow = ThisWorkBook.Worksheets("email handler").Range("F8")
    
    Select Case LastRow

        Case "5 weeks"
          posi = 5
          
        Case "10 weeks"
            posi = 10
        
        Case "13 weeks"
          posi = 13
        
        Case "26 weeks"
          posi = 26

    End Select
    
    'range of header and roster lines from dcam sheets
    Set getRange = ThisWorkBook.Worksheets(dcam).Range("A4:Q" & CStr(posi + 5))

End Function

Sub Btn_ClearCells()

    Range("A2:E600").Clear

End Sub

