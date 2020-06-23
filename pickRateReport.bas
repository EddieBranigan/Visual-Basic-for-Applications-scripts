Attribute VB_Name = "Main"
Option Explicit
'Eddie Branigan

    '22/06/2020 - built method and function layout
    '23/06/2020 - creating methods, testing, bug fixing

Global objSessions          As Object
Global objSystem            As Object
Global Denver               As Object

Sub runMacro()
    Set objSystem = CreateObject("BlueZone.System")
    Set objSessions = objSystem.Sessions
    Set Denver = objSystem.ActiveSession
    
    navDenverMain
    wait
    getAllReports
    wait
End Sub

Sub wait()
    Denver.screen.WaitHostQuiet (1)
End Sub

Sub navDenverMain()
    'check for main screen
    Denver.screen.SendKeys ("<RESET>")
    wait
    Denver.screen.SendKeys ("<PF3><PF3><PF3><PF3><PF3><PF3>")
    wait
    Denver.screen.SendKeys ("denv<Enter>")
    wait
    Denver.screen.SendKeys ("5<Enter>")
    wait
End Sub

Sub navSUBL()
    'check for mainscreen
    Denver.screen.putString "SUBL", 3, 41
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.putString "X", 3, 77
    wait
    Denver.screen.putString "LF25", 4, 75
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
End Sub

Sub navLF25(sTime As String)
    'check for subl
    Dim staStr As String
    Dim finStr As String
    Dim dateCell As Date
    
    dateCell = ThisWorkbook.Worksheets("LF25").Range("F11")
    dateCell = Format(dateCell, "ddmmyy")
    staStr = Format(sTime, "hhmm")
    finStr = TimeValue(sTime) + TimeValue("01:00")
    finStr = Format(finStr, "hhmm")
    Denver.screen.putString dateCell, 5, 40
    Denver.screen.putString Left(staStr, 2), 7, 22
    Denver.screen.putString Right(staStr, 2), 7, 27
    Denver.screen.putString Left(finStr, 2), 7, 32
    Denver.screen.putString Right(finStr, 2), 7, 37
    Denver.screen.putString "Y", 7, 68
    Denver.screen.putString "N", 8, 68
    Denver.screen.putString "N", 9, 68
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.SendKeys ("<PF3>")
    wait
    Denver.screen.SendKeys ("<PF24>")
    wait
End Sub

Function setRepID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LF25")
    Dim today As String
    
    today = Format(Now(), "hh:mm")
    setRepID = ws.Cells(11, 6).Text & "  " & today & " LF25RPT1"
End Function

Sub navMIMX()
    'check for mainscreen
    Denver.screen.putString "MIMX", 3, 41
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.putString "PRTQ", 3, 60
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.putString "X", 22, 27
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
End Sub

Sub findReport(repID As String)
    'check if screen says job queue
    Dim checkMoreStr As String
    Dim row As Integer
    Dim repNum As String
    checkMoreStr = Denver.screen.area(22, 76, 22, 80)
    Do Until checkMoreStr <> "/MORE"
        Denver.screen.SendKeys ("<PF5>")
        wait
        checkMoreStr = Denver.screen.area(22, 76, 22, 80)
    Loop
    For row = 19 To 3 Step -1
        If Denver.screen.area(row, 15, row, 38) = repID Then
               repNum = Denver.screen.area(row, 5, row, 6)
               Denver.screen.putString repNum, 22, 30
               wait
               Denver.screen.SendKeys ("<Enter>")
               wait
        End If
    Next row
End Sub

Sub getReport(row As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LF25")
    
    'check screen for printer selection
    Denver.screen.putString "view", 22, 31
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    ws.Range("C" & row) = Denver.screen.area(9, 48, 9, 51)
    wait
    Denver.screen.SendKeys ("<PF11>")
    wait
    ws.Range("D" & row) = Denver.screen.area(9, 48, 9, 53)
    wait

End Sub

Sub getAllReports()
    Dim ws As Worksheet
    Dim x As Integer
    Dim repID As String
    Set ws = ThisWorkbook.Worksheets("LF25")
    For x = 2 To 17
        navDenverMain
        navSUBL
        navLF25 (ws.Cells(x, 1).Text) 'type mismatch
        repID = setRepID
        navDenverMain
        navMIMX
        findReport (repID)
        getReport (CStr(x))
    Next x
    navDenverMain
    MsgBox ("All reports successfully created.")
End Sub

Sub testingFormat()
    Dim var As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LF25")
    var = ws.Cells(2, 1).Text
    MsgBox (var)
    var = Format(var, "hhmm")
    MsgBox (var)
End Sub

Sub testingTimeAddition()
    Dim answer As String
    answer = TimeValue("07:00") + TimeValue("01:00")
    answer = Format(answer, "hhmm")
    MsgBox (answer)
End Sub

Sub testingNowFunc()
    Dim today As String
    today = Format(Now(), "hhmm")
    MsgBox (today)
End Sub
