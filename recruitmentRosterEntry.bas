Attribute VB_Name = "Main"
Option Explicit
'Eddie Branigan 03/07/2020

Global objSessions          As Object
Global objSystem            As Object
Global Denver               As Object

Sub runMacro()

    Set objSystem = CreateObject("BlueZone.System")
    Set objSessions = objSystem.Sessions
    Set Denver = objSystem.ActiveSession

    Dim row As Integer
    Dim rowCount As Integer
    Dim dcam As String
    Dim startTime As String
    Dim endTime As String
    Dim shiftLength As String
    Dim givenDate As String

    rowCount = getRowCount
    givenDate = getDate()
    shiftLength = ThisWorkbook.Worksheets(1).Cells(4, 8).Text
    
    For row = 2 To rowCount
    
        If ThisWorkbook.Worksheets(1).Cells(row, 1) = "" Then
            GoTo EndOfLoop
        End If
        
        If ThisWorkbook.Worksheets(1).Cells(row, 1) = "NEW" Then
            GoTo InvalidDcam:
        End If
    
        dcam = getCellValue(row, 1)
        startTime = getCellValue(row, 3)
        endTime = getEndTime(row, 3)
    
        navMENU
        navLEMA
        
        If checkDCAM(dcam) Then
            Denver.screen.SendKeys ("<PF7>")
            wait
            Denver.screen.putstring givenDate, 3, 52
            wait
            Denver.screen.SendKeys "<Enter>"
            wait
            updateOrAdd
            enterWorkingHours startTime, endTime, shiftLength
            ThisWorkbook.Worksheets(1).Cells(row, 4) = "Complete"
            ThisWorkbook.Worksheets(1).Cells(row, 4).Interior.ColorIndex = 4
            
        Else
InvalidDcam:
            ThisWorkbook.Worksheets(1).Cells(row, 4) = "Invalid DCAM"
            ThisWorkbook.Worksheets(1).Cells(row, 4).Interior.ColorIndex = 3
        End If
    
EndOfLoop:
    Next row

End Sub

Sub wait()

    Denver.screen.WaitHostQuiet (1)
    
End Sub

Sub navMENU()

    Dim screenCheck As String

    Denver.screen.SendKeys ("<RESET>")
    wait
    Denver.screen.SendKeys ("<PF3><PF3><PF3><PF3>")
    wait
    Denver.screen.SendKeys ("denv<Enter>")
    wait
    Denver.screen.SendKeys ("5<Enter>")
    wait
    
    screenCheck = Denver.screen.Area(24, 6, 24, 9)
        If screenCheck <> "MENU" Then
        MsgBox ("Couldn't reach Denver MENU screen. Is Denver open and are you on the menu screen?")
        End
    End If

End Sub

Sub navLEMA()

    Denver.screen.putstring "LEMA", 3, 41
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait

End Sub

Function getRowCount() As Integer

    With ActiveSheet
        getRowCount = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    
End Function

Function getNumbersOnly(strSource As String) As String

    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57:
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    
    getNumbersOnly = strResult

End Function

Function getCellValue(row As Integer, col As Integer) As String
    
    Dim cellValue As String
    
    cellValue = ThisWorkbook.Worksheets(1).Cells(row, col).Text
    
    getCellValue = getNumbersOnly(cellValue)
    
End Function

Sub updateOrAdd()

    Dim checkScreen As String
    
    checkScreen = Denver.screen.Area(24, 40, 24, 41)
    
    If checkScreen = "96" Then
    
        Denver.screen.putstring "ADD", 3, 73
        wait
        Denver.screen.SendKeys ("<Enter>")
        wait
        
    ElseIf checkScreen = "71" Then
    
        Denver.screen.putstring "Update", 3, 73
        wait
        Denver.screen.SendKeys ("<Enter>")
        wait
        
    End If
    
End Sub

Function getEndTime(row As Integer, col As Integer) As String
    
    Dim startTime As String
    Dim givenShift As Integer
    Dim givenTime As Date
    
    startTime = ThisWorkbook.Worksheets(1).Cells(row, col).Value
    givenShift = ThisWorkbook.Worksheets(1).Cells(4, 8).Value
    givenTime = CDate(startTime)
    givenTime = DateAdd("h", givenShift, givenTime)

    getEndTime = Format(givenTime, "hhmm")

End Function

Function getDate() As String

    Dim givenDate As String
    givenDate = ThisWorkbook.Worksheets(1).Cells(3, 8).Text
    getDate = getNumbersOnly(givenDate)
    MsgBox (getDate)
    
End Function

Function checkDCAM(dcam As String) As Boolean

    Dim errorCheck As String
    checkDCAM = False
    
    Denver.screen.putstring dcam, 2, 73
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    errorCheck = Denver.screen.Area(24, 40, 24, 41)
    If errorCheck <> "98" Then
        checkDCAM = True
    End If
    
End Function

Sub enterWorkingHours(startTime As String, finishTime As String, shiftLength As String)

    Dim givenDay As String
    givenDay = ThisWorkbook.Worksheets(1).Cells(2, 8).Text
    MsgBox (givenDay)

    Select Case givenDay
    
        Case "Sunday"
            Denver.screen.putstring startTime, 6, 28
            wait
            Denver.screen.putstring finishTime, 7, 28
            wait
            Denver.screen.putstring shiftLength, 8, 30
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait

        Case "Monday"
            Denver.screen.putstring startTime, 6, 36
            wait
            Denver.screen.putstring finishTime, 7, 36
            wait
            Denver.screen.putstring shiftLength, 8, 38
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait

        Case "Tuesday"
            Denver.screen.putstring startTime, 6, 44
            wait
            Denver.screen.putstring finishTime, 7, 44
            wait
            Denver.screen.putstring shiftLength, 8, 46
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait

        Case "Wednesday"
            Denver.screen.putstring startTime, 6, 52
            wait
            Denver.screen.putstring finishTime, 7, 52
            wait
            Denver.screen.putstring shiftLength, 8, 54
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait


        Case "Thursday"
            Denver.screen.putstring startTime, 6, 60
            wait
            Denver.screen.putstring finishTime, 7, 60
            wait
            Denver.screen.putstring shiftLength, 8, 62
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait

        Case "Friday"
            Denver.screen.putstring startTime, 6, 68
            wait
            Denver.screen.putstring finishTime, 7, 68
            wait
            Denver.screen.putstring shiftLength, 8, 70
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait

        Case "Saturday"
            Denver.screen.putstring startTime, 6, 76
            wait
            Denver.screen.putstring finishTime, 7, 76
            wait
            Denver.screen.putstring shiftLength, 8, 78
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait

    End Select

End Sub

Sub clearScreen()

    Dim rowCount As Integer
    
    With ActiveSheet
        rowCount = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    
    ActiveSheet.Range("A2:D" & rowCount).Clear
    
End Sub
Sub testgetDate()

    Dim givenDate As String
    givenDate = getDate()
    MsgBox (givenDate)
    'Working
    
End Sub

Sub testShiftLength()

    MsgBox (ThisWorkbook.Worksheets(1).Cells(4, 8).Text)
    'Working

End Sub

Sub testCDate()

    Dim test As String
    Dim dateTest As String
    
    test = "09:00"
    dateTest = CDate(test)
    MsgBox (dateTest)
    'working
    
End Sub
