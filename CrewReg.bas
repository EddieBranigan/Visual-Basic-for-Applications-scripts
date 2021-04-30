Attribute VB_Name = "CrewReg"
Option Explicit
'Eddie Branigan 2020

Public Const ThisWb As String = "Crew Reg Macro v1.05.xlsm"

Public Sessions As Object
Public System As Object
Public Denver As Object

Public Sub getCrewReg()
    'Fetches all the rosters from the four roster workbooks
    'and compiles them into a crew reg worksheet along with
    'the skills of assemblers. Screen updating and a
    'loading bar is displayed for a more accessible user experience
    
    Dim i As Integer
    Application.ScreenUpdating = False
    ActiveWindow.Visible = True
    clearWS
    getCrewRoster ("Roster1")
    progress (20)
    getCrewRoster ("Roster2")
    progress (40)
    getCrewRoster ("Roster3")
    progress (60)
    getCrewRoster ("Roster4")
    progress (80)
    getTempAndSkill
    progress (99)
    makeTable
    sortByTimes
    FormatTables
    Unload ProgBar
    Application.ScreenUpdating = True
    MsgBox ("Finished running macro")
End Sub

Public Sub getCRworked()
    'Connects to Denver through the bluezone object(activex).
    'If it fails to connect to reach the appropriate screen,
    'it gives an error message. If it connects it displays a
    'progrees bar as it fetches clock times for the specified days
    'on denver
    
    Set System = CreateObject("BlueZone.System")
    Set Sessions = System.Sessions
    Set Denver = System.ActiveSession
    Dim xs As Integer
    
    For xs = 1 To Sessions.Count
        If InStr(LCase(Sessions.Item(xs).name), "Denver") Then
            Set Denver = Sessions.Item(xs)
        End If
    Next xs
                            
    navMAIN
    If Denver.screen.area(24, 6, 24, 9) <> "MENU" Then
        MsgBox ("Cannot reach Denver Main Menu. Make sure you are logged on to Denver and on the MAIN screen.")
    Else
        getCRTimes
        flagCells
        Unload denvProgBar
        MsgBox ("Finished running macro")
    End If
End Sub

Private Sub FormatTables()
    'unprotects the crew reg worksheet and creates,
    'headers, tables and numerous formats
    
    Workbooks(ThisWb).Worksheets("Crew Reg").Unprotect
    formatSkillMix
    formatTimeCols
    formatDCAMS
    formatNameCols
    formatDateCell
    formatSusTime
    formatRosNo
    formatDate
    formatComments
    lockCells
End Sub

Private Sub wait()
    'instructs the denver object to wait until it
    'recieves a ready instruction (displayed at the bottom of the screen)
    
    Denver.screen.WaitHostQuiet (1)
End Sub

Private Sub navMAIN()
    'instructions to navigate Denver to the MAIN screen
    
    Denver.screen.SendKeys ("<RESET>")
    wait
    Denver.screen.SendKeys ("<PF3><PF3><PF3><PF3>")
    wait
    Denver.screen.SendKeys ("denv")
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.SendKeys ("5")
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
End Sub

Private Sub navFreezeMain()
    'instructions to navigate Denver to the Freezer MAIN screen
    
    Denver.screen.SendKeys ("<RESET>")
    wait
    Denver.screen.SendKeys ("<PF3><PF3><PF3><PF3>")
    wait
    Denver.screen.SendKeys ("denv")
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.SendKeys ("1")
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
End Sub

Private Sub navLAQA()
    'instructions to navigate Denver to the LAQA screen
    
    Denver.screen.SendKeys ("LAQA")
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
End Sub

Private Sub getCRTimes()
    Dim rosWS As Worksheet
    Dim timeString As String
    Dim cRow As Integer
    Dim lastrow As Integer
    Dim lunchAmounts As Integer
    Dim dcam As String
    Dim watcher As Boolean
    Dim denScrnRow As Integer
    Dim assignmentType As String
    Dim overTimes As Integer
    Dim dcamAdd As Range
    Dim totalLines As Integer
    timeString = getCRTimeString
    Set rosWS = ThisWorkbook.Worksheets("Crew Reg")
    lastrow = rosWS.Cells(Rows.Count, "A").End(xlUp).Row
    navLAQA
    ThisWorkbook.Sheets("Crew Reg").Unprotect
    
    'loops through all lines of Crew Reg. sheet
    For cRow = 3 To lastrow
        totalLines = (lastrow - 3) / 100
        progress2 (cRow / totalLines)
        lunchAmounts = 0
        overTimes = 0

        If Not IsEmpty(rosWS.Cells(cRow, 16)) And Not IsEmpty(rosWS.Cells(cRow, 16)) Then
        '
        Else
        'loops through lines of denver LAQA screen
        dcam = rosWS.Cells(cRow, 4).Text
        Denver.screen.PutString dcam, 2, 68
        wait
        Denver.screen.PutString Mid(timeString, 1, 2), 3, 68
        Denver.screen.PutString Mid(timeString, 3, 2), 3, 73
        Denver.screen.PutString Mid(timeString, 5, 2), 3, 78
        wait
        Denver.screen.SendKeys ("<Enter>")
        wait
            For denScrnRow = 9 To 18
                assignmentType = ""
                assignmentType = Denver.screen.area(denScrnRow, 13, denScrnRow, 18)
                
                'records different assignment types
                Select Case assignmentType
                    Case "OUTDAY"
                        rosWS.Cells(cRow, 16) = Denver.screen.area(denScrnRow, 28, denScrnRow, 32)
                    Case "IN DAY"
                        rosWS.Cells(cRow, 15) = Denver.screen.area(denScrnRow, 28, denScrnRow, 32)
                    Case "LUNCH "
                        lunchAmounts = lunchAmounts + 1
                    Case "OT BRK"
                        overTimes = overTimes + 1
                End Select
                
                'puts suspend time in col
                If Denver.screen.area(24, 40, 24, 41) = "95" Then
                    If Denver.screen.area(20, 72, 20, 73) <> "  " Then
                        rosWS.Cells(cRow, 17) = Denver.screen.area(20, 72, 20, 73)
                    Else
                        rosWS.Cells(cRow, 17) = "0"
                    End If
                End If
                    
                'allows to browse all screens
                If denScrnRow = 18 And Denver.screen.area(24, 40, 24, 41) = "97" Then
                    denScrnRow = 8
                    Denver.screen.SendKeys ("<PF2>")
                    wait
                End If
                
                'appends lunchtimes to suspend col
                If denScrnRow = 18 And Denver.screen.area(24, 40, 24, 41) = "95" Then
                    If rosWS.Cells(cRow, 15) <> "" Then
                        rosWS.Cells(cRow, 17) = overTimes & " | " & lunchAmounts
                    Else
                        rosWS.Cells(cRow, 17) = ""
                    End If
                End If
            Next denScrnRow
        End If
    
    Next cRow
    
    navFreezeMain
    navLAQA
    For cRow = 3 To lastrow
        progress2 (cRow / totalLines)
        If rosWS.Cells(cRow, 16).Interior.Color = RGB(252, 107, 3) Then
            dcam = rosWS.Cells(cRow, 4).Text
            Denver.screen.PutString dcam, 2, 68
            wait
            Denver.screen.PutString Mid(timeString, 1, 2), 3, 68
            Denver.screen.PutString Mid(timeString, 3, 2), 3, 73
            Denver.screen.PutString Mid(timeString, 5, 2), 3, 78
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait
            
            For denScrnRow = 9 To 18
                assignmentType = ""
                assignmentType = Denver.screen.area(denScrnRow, 13, denScrnRow, 18)
                
                Select Case assignmentType
                    Case "OUTDAY"
                        rosWS.Cells(cRow, 16) = Denver.screen.area(denScrnRow, 28, denScrnRow, 32)
                        rosWS.Cells(cRow, 1).Interior.ColorIndex = 4
                        rosWS.Cells(cRow, 2).Interior.ColorIndex = 4
                        rosWS.Cells(cRow, 3).Interior.ColorIndex = 4
                        rosWS.Cells(cRow, 4).Interior.ColorIndex = 4
                    Case "IN DAY"
                        'Not needed at the moment
                    Case "LUNCH "
                        'Not needed at the moment
                End Select
                
                'allows to browse all screens
                If denScrnRow = 18 And Denver.screen.area(24, 40, 24, 41) = "97" Then
                    denScrnRow = 8
                    Denver.screen.SendKeys ("<PF2>")
                    wait
                End If
            Next denScrnRow
            
        End If
    Next cRow
    
    navMAIN
    ThisWorkbook.Worksheets("Crew Reg").Protect AllowFormattingCells:=True
End Sub

Private Sub getCrewRoster(rosterName As String)
    Dim currentWB As Workbook
    Dim stringFileLoc As String
    Dim rosterWB As Workbook
    Dim i As Integer
    'stringFileLoc = ""
    stringFileLoc = ""
    Set rosterWB = Workbooks.Open(stringFileLoc & rosterName & ".xls", _
    UpdateLinks:=False, ReadOnly:=True)
    Dim wsCount As Integer
    wsCount = rosterWB.Worksheets.Count
    For i = 1 To wsCount
        If regExNameTester(rosterWB.Worksheets(i).name) Then
            If rosterWB.Worksheets(i).Range("B7") <> "" _
            And rosterWB.Worksheets(i).Range("A5") <> "" _
            And rosterWB.Worksheets(i).Range("B5") <> "" Then
                getCrewRosterInfo rosterWB, rosterWB.Worksheets(i).name
            End If
        End If
    Next i
    rosterWB.Saved = True
    rosterWB.Close
End Sub

Private Sub getCrewRosterInfo(rosterWB As Workbook, sheetName As String)
    'looks through a specified sheet contained in one of the roster
    'workbooks and searches for the date given on the plannin worksheet.
    'When found it checks the day and using a select case function,
    'sorts the roster start and finish times into their appropriate cells on
    'the crew reg. worksheet.
    
    Dim cr As Integer
    Dim rosters As Worksheet
    Dim day As String
    Dim dateRow As Integer
    Dim foundCell As Range
    Set rosters = Workbooks(ThisWb).Worksheets("Crew Reg")
    day = Format(Workbooks(ThisWb).Worksheets("Planning").Range("E3").Text, "dddd")
    cr = Workbooks(ThisWb).Worksheets("Crew Reg").Cells(Rows.Count, "A").End(xlUp).Row + 1 'cr is current row
    Set foundCell = rosterWB.Worksheets(sheetName).Range("B:B").Find(What:=Workbooks(ThisWb).Worksheets("Directory").Range("E2").Text)
    dateRow = foundCell.Row
    rosters.Cells(cr, 1) = rosterWB.Worksheets(sheetName).Range("B4") 'roster number
    rosters.Cells(cr, 4) = rosterWB.Worksheets(sheetName).Range("B7") 'Dcam
    rosters.Cells(cr, 3) = rosterWB.Worksheets(sheetName).Range("B5") 'Name
    rosters.Cells(cr, 2) = rosterWB.Worksheets(sheetName).Range("A5") 'Surname
    
    Select Case day
        Case "Sunday"
            rosters.Cells(cr, 13) = rosterWB.Worksheets(sheetName).Cells(dateRow, 3) 'start time
            rosters.Cells(cr, 14) = rosterWB.Worksheets(sheetName).Cells(dateRow, 4) 'finish time
            If rosterWB.Worksheets(sheetName).Cells(dateRow, 3).Interior.ColorIndex = 6 Then
                'rosters.Cells(cr, 1).Interior.ColorIndex = 4
                'rosters.Cells(cr, 2).Interior.ColorIndex = 4
                'rosters.Cells(cr, 3).Interior.ColorIndex = 4
                'rosters.Cells(cr, 4).Interior.ColorIndex = 4
            End If
        Case "Monday"
            rosters.Cells(cr, 13) = rosterWB.Worksheets(sheetName).Cells(dateRow, 8) 'start time
            rosters.Cells(cr, 14) = rosterWB.Worksheets(sheetName).Cells(dateRow, 9) 'finish time
            If rosterWB.Worksheets(sheetName).Cells(dateRow, 8).Interior.ColorIndex = 6 Then
                'rosters.Cells(cr, 1).Interior.ColorIndex = 4
                'rosters.Cells(cr, 2).Interior.ColorIndex = 4
                'rosters.Cells(cr, 3).Interior.ColorIndex = 4
                'rosters.Cells(cr, 4).Interior.ColorIndex = 4
            End If
        Case "Tuesday"
            rosters.Cells(cr, 13) = rosterWB.Worksheets(sheetName).Cells(dateRow, 13) 'start time
            rosters.Cells(cr, 14) = rosterWB.Worksheets(sheetName).Cells(dateRow, 14)  'finish time
            If rosterWB.Worksheets(sheetName).Cells(dateRow, 13).Interior.ColorIndex = 6 Then
                'rosters.Cells(cr, 1).Interior.ColorIndex = 4
                'rosters.Cells(cr, 2).Interior.ColorIndex = 4
                'rosters.Cells(cr, 3).Interior.ColorIndex = 4
                'rosters.Cells(cr, 4).Interior.ColorIndex = 4
            End If
        Case "Wednesday"
            rosters.Cells(cr, 13) = rosterWB.Worksheets(sheetName).Cells(dateRow, 18) 'start time
            rosters.Cells(cr, 14) = rosterWB.Worksheets(sheetName).Cells(dateRow, 19) 'finish time
            If rosterWB.Worksheets(sheetName).Cells(dateRow, 18).Interior.ColorIndex = 6 Then
                'rosters.Cells(cr, 1).Interior.ColorIndex = 4
                'rosters.Cells(cr, 2).Interior.ColorIndex = 4
                'rosters.Cells(cr, 3).Interior.ColorIndex = 4
                'rosters.Cells(cr, 4).Interior.ColorIndex = 4
            End If
        Case "Thursday"
            rosters.Cells(cr, 13) = rosterWB.Worksheets(sheetName).Cells(dateRow, 23) 'start time
            rosters.Cells(cr, 14) = rosterWB.Worksheets(sheetName).Cells(dateRow, 24) 'finish time
            If rosterWB.Worksheets(sheetName).Cells(dateRow, 23).Interior.ColorIndex = 6 Then
                'rosters.Cells(cr, 1).Interior.ColorIndex = 4
                'rosters.Cells(cr, 2).Interior.ColorIndex = 4
                'rosters.Cells(cr, 3).Interior.ColorIndex = 4
                'rosters.Cells(cr, 4).Interior.ColorIndex = 4
            End If
        Case "Friday"
            rosters.Cells(cr, 13) = rosterWB.Worksheets(sheetName).Cells(dateRow, 28) 'start time
            rosters.Cells(cr, 14) = rosterWB.Worksheets(sheetName).Cells(dateRow, 29) 'finish time
            If rosterWB.Worksheets(sheetName).Cells(dateRow, 28).Interior.ColorIndex = 6 Then
                'rosters.Cells(cr, 1).Interior.ColorIndex = 4
                'rosters.Cells(cr, 2).Interior.ColorIndex = 4
                'rosters.Cells(cr, 3).Interior.ColorIndex = 4
                'rosters.Cells(cr, 4).Interior.ColorIndex = 4
            End If
        Case "Saturday"
            rosters.Cells(cr, 13) = rosterWB.Worksheets(sheetName).Cells(dateRow, 33) 'start time
            rosters.Cells(cr, 14) = rosterWB.Worksheets(sheetName).Cells(dateRow, 34) 'finish time
            If rosterWB.Worksheets(sheetName).Cells(dateRow, 33).Interior.ColorIndex = 6 Then
                'rosters.Cells(cr, 1).Interior.ColorIndex = 4
                'rosters.Cells(cr, 2).Interior.ColorIndex = 4
                'rosters.Cells(cr, 3).Interior.ColorIndex = 4
                'rosters.Cells(cr, 4).Interior.ColorIndex = 4
            End If
   End Select
End Sub

Private Sub getTempAndSkill()
    'Checks the Name File in the roster folder for the skills of assemblers
    'and appends them to their appropriate areas.
    
    Dim rosters As Worksheet
    Dim rosterWB As Workbook
    Dim stringFileLoc As String
    Dim lastrow As Integer
    Dim i As Integer
    Dim c As Range
    Dim dateRow As Integer
    Dim Line As Integer
    Dim recruit As String
    'stringFileLoc = "C:\Users\IEE12367699\OneDrive - Tesco\office work\testing files\"
    stringFileLoc = "\\global.tesco.org\dfsroot\IE\Distribution\Ballymun\Planning\Warehouse Planning\Rosters\Individual\"
    Set rosters = Workbooks(ThisWb).Worksheets("Crew Reg")
    Set rosterWB = Workbooks.Open(stringFileLoc & "Names File" & ".xls", _
    UpdateLinks:=False, ReadOnly:=True)
    lastrow = Workbooks(ThisWb).Worksheets("Crew Reg").Cells(Rows.Count, "A").End(xlUp).Row
    For i = 3 To lastrow
        With rosterWB.Worksheets("Namepage").Range("K:K")
            Set c = .Find(rosters.Cells(i, 4).Text, LookIn:=xlValues)
            If Not c Is Nothing Then
                dateRow = c.Row
                Do
                    c.Value = 5
                    Set c = .FindNext(c)
                Loop While Not c Is Nothing
            End If
        End With
    'unloading - 5
    rosters.Cells(i, 5) = rosterWB.Worksheets("Namepage").Cells(dateRow, 5)
    'checking - 6
    rosters.Cells(i, 6) = rosterWB.Worksheets("Namepage").Cells(dateRow, 6)
    'loading - 7
    rosters.Cells(i, 7) = rosterWB.Worksheets("Namepage").Cells(dateRow, 7)
    'paperless - 8
    rosters.Cells(i, 8) = rosterWB.Worksheets("Namepage").Cells(dateRow, 8)
    'battery - 13
    rosters.Cells(i, 9) = rosterWB.Worksheets("Namepage").Cells(dateRow, 13)
    'layermaster - 16
    rosters.Cells(i, 10) = rosterWB.Worksheets("Namepage").Cells(dateRow, 16)
    'd/d - 19
    rosters.Cells(i, 11) = rosterWB.Worksheets("Namepage").Cells(dateRow, 19)
    'topping - 20
    rosters.Cells(i, 12) = rosterWB.Worksheets("Namepage").Cells(dateRow, 20)
    Next i
    For Line = 250 To 600
        recruit = rosterWB.Worksheets("Namepage").Cells(Line, 5).Text
        Select Case recruit
            Case "Temple"
                rosters.Cells(lastrow, 1) = "TEMP"
                rosters.Cells(lastrow, 2) = rosterWB.Worksheets("Namepage").Cells(Line, 2).Text
                rosters.Cells(lastrow, 3) = rosterWB.Worksheets("Namepage").Cells(Line, 3).Text
                rosters.Cells(lastrow, 4) = rosterWB.Worksheets("Namepage").Cells(Line, 4).Text
            Case "Flex"
                rosters.Cells(lastrow, 1) = "FLEX"
                rosters.Cells(lastrow, 2) = rosterWB.Worksheets("Namepage").Cells(Line, 2).Text
                rosters.Cells(lastrow, 3) = rosterWB.Worksheets("Namepage").Cells(Line, 3).Text
                rosters.Cells(lastrow, 4) = rosterWB.Worksheets("Namepage").Cells(Line, 4).Text
        End Select
        lastrow = Workbooks(ThisWb).Worksheets("Crew Reg").Cells(Rows.Count, "A").End(xlUp).Row + 1
    Next Line
    rosterWB.Saved = True
    rosterWB.Close
End Sub

Private Function regExNameTester(workSheetName As String)
    'Checks the name of the worksheet to see if it is between
    'one and four digits long
    
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
      .Pattern = "^[0-9 ]+$"
    End With
    regExNameTester = RegEx.test(workSheetName)
End Function

Private Function getCRTimeString() As String
    'converts a string that is a date to no a six digit long
    'value for use with Denver
    
    Dim dt As Date
    dt = Workbooks(ThisWb).Worksheets("Crew Reg").Range("Q1")
    getCRTimeString = Format(dt, "ddmmyy")
End Function

'Formatting------------------------------------------------------------
Private Sub clearWS()
    'Converts the contents of the sheet from a table to range class
    'to be cleared of all data and formats
    
    ConvertTablesToRange
    Workbooks(ThisWb).Sheets("Crew Reg").Unprotect
    Dim lastrow As String
    Dim wrkSht As Worksheet
    Set wrkSht = Workbooks(ThisWb).Worksheets("Crew Reg")
    lastrow = wrkSht.Cells(Rows.Count, "A").End(xlUp).Row
    If wrkSht.Range("A3") = "" Then
        lastrow = wrkSht.Cells(Rows.Count, "A").End(xlUp).Row + 1
        wrkSht.Range("A3:R" & lastrow).ClearContents
        wrkSht.Range("A3:R" & lastrow).ClearFormats
    Else
        wrkSht.Range("A3:R" & lastrow).ClearContents
        wrkSht.Range("A3:R" & lastrow).ClearFormats
    End If
End Sub

Private Sub ConvertTablesToRange()
    Dim wks As Worksheet, objList As ListObject
    Set wks = Workbooks(ThisWb).Worksheets("Crew Reg")
    For Each objList In wks.ListObjects
        objList.Unlist
    Next objList
End Sub

Private Function TableExists() As Boolean
    TableExists = False
    On Error GoTo Skip
    If ActiveSheet.ListObjects("Crew Reg").name = "crewRegTable" Then TableExists = True            'CHECK THIS WORKS WITHOUT ACTIVESHEET
    
Skip:
    On Error GoTo 0
End Function

Private Sub makeTable()
    Dim lastrow As Integer
    lastrow = Workbooks(ThisWb).Worksheets("Crew Reg").Cells(Rows.Count, "A").End(xlUp).Row
    With Workbooks(ThisWb).Worksheets("Crew Reg").Range("A3:R" & CStr(lastrow))
        Workbooks(ThisWb).Worksheets("Crew Reg").ListObjects.Add(xlSrcRange, _
        Workbooks(ThisWb).Worksheets("Crew Reg").Range("$A$2:$R$" & CStr(lastrow)), , xlYes).name = _
            "crewRegTable"
    End With
End Sub

Private Sub sortByTimes()
    'Sorts the rows of the workbook by the starting times of the assemblers                         'CHECK THIS WORKS WITHOUT ACTIVEWORKBOOK
    
    ActiveWorkbook.Worksheets("Crew Reg").ListObjects("crewRegTable").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("Crew Reg").ListObjects("crewRegTable").Sort. _
        SortFields.Add Key:=Range("crewRegTable[[#All],[Roster Start]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Crew Reg").ListObjects("crewRegTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub formatSkillMix()
    'Formats the skill mix section of the worksheet to look
    'similar to the printed version of the Crew Reg
    
    With Workbooks(ThisWb).Worksheets("Crew Reg").Columns("E:L")
        .HorizontalAlignment = xlCenter
        .ColumnWidth = 2.3
        .NumberFormat = "General"
    End With
    Dim lastrow As String
    lastrow = Workbooks(ThisWb).Worksheets("Crew Reg").Cells(Rows.Count, "A").End(xlUp).Row
    Sheets("Crew Reg").Activate
    ActiveSheet.Range("E3:L" & lastrow).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
       .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
       .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
       .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
       .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlHairline
    End With
End Sub

Private Sub formatTimeCols()
    With Workbooks(ThisWb).Worksheets("Crew Reg").Columns("M:P")
        .HorizontalAlignment = xlCenter
        .ColumnWidth = 6.5
        .NumberFormat = "h:mm"
    End With
    Dim lastrow As Integer
    lastrow = Workbooks(ThisWb).Worksheets("Crew Reg").Cells(Rows.Count, "A").End(xlUp).Row
    Sheets("Crew Reg").Activate
    ActiveSheet.Range("N3:N" & lastrow).Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlSlantDashDot
        .ColorIndex = 46
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End Sub

Private Sub formatDCAMS()
    With Workbooks(ThisWb).Worksheets("Crew Reg").Columns("D:D")
        .HorizontalAlignment = xlCenter
        .ColumnWidth = 6
        .NumberFormat = "General"
    End With
End Sub

Private Sub formatNameCols()
    With Workbooks(ThisWb).Worksheets("Crew Reg").Columns("B:C")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ColumnWidth = 20
    End With
End Sub

Private Sub formatDateCell()
    With Workbooks(ThisWb).Worksheets("Crew Reg").Range("Q1:R1")
        .ColumnWidth = 10
    End With
End Sub

Private Sub formatComments()
    With Workbooks(ThisWb).Worksheets("Crew Reg").Range("R:R")
        .ColumnWidth = 40
    End With
End Sub

Private Sub formatRosNo()
    With Workbooks(ThisWb).Worksheets("Crew Reg").Range("A:A")
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub formatSusTime()
    With Workbooks(ThisWb).Worksheets("Crew Reg").Range("Q:Q")
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub formatDate()
    Workbooks(ThisWb).Worksheets("Planning").Range("E3").Copy _
    Workbooks(ThisWb).Worksheets("Crew Reg").Range("Q1:R1")
    Workbooks(ThisWb).Worksheets("Crew Reg").Range("Q1").Select
    With Selection.Font
        .name = "Calibri"
        .Size = 20
        .Shadow = True
    End With
End Sub
Private Sub CopyItOver()
    'Makes a copy of this workbook, deletes all sheets but the Crew Reg and
    'saves a copy in the Crew Reg folder in the warehouse directory.
    'It is named after the day the crew reg is generated for. ex: Crew Reg "01-01-2010.xlsm"
    
    Application.ScreenUpdating = False
    ActiveWindow.Visible = True
    Dim targetWB As Workbook
    Dim FName As String
    Dim fpath As String
    Dim strTempFile As String
    Dim ws As Worksheet
    Dim dt As Date
    Dim WkSht As Worksheet
    dt = Workbooks(ThisWb).Worksheets("Crew Reg").Range("Q1")
    fpath = "X:\Warehouse\Crew Reg\2020\"
    FName = "Crew Reg " & Format(dt, "dd-mm-yyyy") & ".xlsm"
    On Error Resume Next
        Set ws = Workbooks(ThisWb).Sheets("Crew Reg")
    On Error GoTo 0
    If ws Is Nothing Then
       MsgBox "sheet doesn't exist"
       Exit Sub
    End If
    If Dir(fpath & "\" & FName) = vbNullString Then
        Workbooks(ThisWb).SaveCopyAs fpath & FName
        Application.Workbooks.Open (fpath & "\" & FName)
        For Each WkSht In ActiveWorkbook.Worksheets
            Select Case WkSht.name
            Case "Planning", "Rosters", "Directory"
                Application.DisplayAlerts = False
                WkSht.Delete
            Case Else
             'do nothing
        End Select
        Application.CutCopyMode = False
        Next WkSht
        ActiveWorkbook.Save
        ActiveWorkbook.Close
    Else
         MsgBox "File " & fpath & "\" & FName & " already exists"
    End If
    Application.ScreenUpdating = True
    MsgBox (FName & " successfully sent to warehouse folder")
End Sub

Private Sub printCrewReg()
    'Prints the Crew Reg. sheet out
    Workbooks(ThisWb).Worksheets("CrewReg").printOut
End Sub

Private Sub flagCells()
    'Checks over the actual start and finsh times and colours them orange if they
    'due to start but haven't. Will only check a cell if its start time is after
    'the current time the macro is ran at. Cells are coloured yellow if a worker
    'arrives late
    
    Dim x As Integer
    Dim rosStart As Date
    Dim actStart As Date
    Dim ws As Worksheet
    Dim time As String
    Dim one As Date
    Dim two As Date
    Dim lateTime As Date
    time = Format(Now, "hh:mm")
    Set ws = Workbooks(ThisWb).Worksheets("Crew Reg")
    For x = 3 To ws.Cells(Rows.Count, "A").End(xlUp).Row
        If ws.Cells(x, 15) <> "" Then
            ws.Cells(x, 15).Interior.Color = ws.Cells(x, 14).Interior.ColorIndex
            If ws.Cells(x, 16) = "" Then
                ws.Cells(x, 16).Interior.Color = RGB(252, 107, 3)
            Else
                ws.Cells(x, 16).Interior.Color = ws.Cells(x, 14).Interior.ColorIndex
            End If
        End If
        
        If ws.Cells(x, 13) <> "" And ws.Cells(x, 15) = "" Then
            If IsNumeric(ws.Cells(x, 13)) Then
                If TimeValue(ws.Cells(x, 13).Text) < TimeValue(time) Then
                    ws.Cells(x, 15).Interior.Color = RGB(252, 107, 3)
                    ws.Cells(x, 16).Interior.Color = RGB(252, 107, 3)
                End If
            End If
        End If
        
        lateTime = TimeValue("00:04:00")
        If IsNumeric(ws.Cells(x, 13).Value) And ws.Cells(x, 13).Value <> "" Then
            one = ws.Cells(x, 13).Value
            two = ws.Cells(x, 15).Value
            
            If two - one > lateTime Then
                ws.Cells(x, 15).Interior.Color = RGB(255, 255, 0)
            End If
        End If
    Next x
End Sub

Private Sub lockCells()
    'locks all cells to prevent 'mistakes'
    
    Dim lastrow As Integer
    Dim titleRng As Range
    Dim rosterRng As Range
    Dim editRosterRng As Range
    lastrow = Workbooks(ThisWb).Worksheets("Crew Reg").Cells(Rows.Count, "A").End(xlUp).Row
    Set rosterRng = Workbooks(ThisWb).Worksheets("Crew Reg").Range("A3:N" & lastrow)
    Set titleRng = Workbooks(ThisWb).Worksheets("Crew Reg").Range("A1:R2")
    Set editRosterRng = Workbooks(ThisWb).Worksheets("Crew Reg").Range("O2:R" & lastrow)
    titleRng.Locked = True
    rosterRng.Locked = True
    editRosterRng.Locked = False
    Worksheets("Crew Reg").Protect AllowFormattingCells:=True
End Sub

Sub progress(pctCompl As Single)
    'An update for the form that appears to show a loading bar.
    'As tasks are completed in the macro, the loading bar is filled
    'further. When the value pctComp1 reaches 200 the bar is full
    ProgBar.labelText.Caption = pctCompl & "% Completed"
    ProgBar.bar.Width = pctCompl * 2
    DoEvents
End Sub

Sub progress2(progNumb As Integer)
    'An update for the form that appears to show a loading bar.
    'As tasks are completed in the macro, the loading bar is filled
    'further. When the value progNumb reaches 200 the bar is full
    denvProgBar.text2.Caption = progNumb & "% Completed"
    denvProgBar.dprogbar.Width = progNumb * 2
    DoEvents
End Sub

Sub startProgBar()
    ProgBar.Show
End Sub

Sub startDenverProgBar()
    denvProgBar.Show
End Sub


