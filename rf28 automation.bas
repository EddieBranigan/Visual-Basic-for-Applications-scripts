Attribute VB_Name = "Main"
Option Explicit

Global objSessions          As Object
Global objSystem            As Object
Global Denver               As Object

Sub Bt_TestInputDenver()
    
    Set objSystem = CreateObject("BlueZone.System")
    Set objSessions = objSystem.Sessions
    Set Denver = objSystem.ActiveSession
    
    Dim rowArray() As Variant
    Dim rowCount As String
    Dim row As Integer
    
    rowCount = getRowCount
    
    'loop through all lines in worksheets
    For row = 2 To rowCount
        
        'sets rowArray as empty
        Erase rowArray
        
        'make an array with all cells from selected row
        rowArray = getArrayFromRow(row)
        
        navDenverMain
        navLEMA
        
        If checkDCAM(CStr(rowArray(1))) Then
            wait
            'go to lsma
            Denver.screen.SendKeys ("<PF7>")
            wait
            'checks cell r7 and converts to date
            Denver.screen.putstring convertDateToString, 3, 52
            wait
            Denver.screen.SendKeys "<Enter>"
            wait
            'checks whether roster is entered or needs to be added
            updateOrAdd
            wait
            'writes roster data to appropriate screen areas
            enterDayHours (rowArray)
            ThisWorkbook.Worksheets(1).Range("P" & row) = "Complete"
            ThisWorkbook.Worksheets(1).Range("P" & row).Interior.ColorIndex = 4
            
        Else
            'set status to "Not on File"
            ThisWorkbook.Worksheets(1).Range("P" & row) = "Invalid DCAM"
            ThisWorkbook.Worksheets(1).Range("P" & row).Interior.ColorIndex = 3
        End If
        
    Next row

End Sub

Function AlphaNumericOnly(strSource As String) As String

    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122:
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    
    AlphaNumericOnly = strResult

End Function

Sub wait()

    Denver.screen.WaitHostQuiet (1)
    
End Sub

Sub navDenverMain()

    Denver.screen.SendKeys ("<RESET>")
    wait
    Denver.screen.SendKeys ("<PF3><PF3><PF3><PF3>")
    wait
    Denver.screen.SendKeys ("denv<Enter>")
    wait
    Denver.screen.SendKeys ("5<Enter>")
    wait
    
End Sub

Sub navLEMA()

    Denver.screen.putstring "LEMA", 3, 41
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait

End Sub

Function getRowCount() As String

    With ActiveSheet
        getRowCount = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    
End Function

Function getArrayFromRow(row As Integer) As Variant
        
        Dim y As Integer
        Dim rowArray(15) As Variant
        
        'set array of cells()
        For y = 1 To 15
            With ActiveSheet
                rowArray(y) = AlphaNumericOnly(Cells(row, y).Text)
            End With
        Next y
        
        getArrayFromRow = rowArray
        
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

Function convertDateToString() As String
    
    Dim cellDate As Date

    cellDate = ThisWorkbook.Worksheets(1).Range("R3")
    convertDateToString = Format(cellDate, "ddmmyy")
    
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

Sub enterDayHours(rowArray As Variant)
            
            enterRosterInfo CStr(rowArray(2)), CStr(rowArray(3)), 28
            enterRosterInfo CStr(rowArray(4)), CStr(rowArray(5)), 36
            enterRosterInfo CStr(rowArray(6)), CStr(rowArray(7)), 44
            enterRosterInfo CStr(rowArray(8)), CStr(rowArray(9)), 52
            enterRosterInfo CStr(rowArray(10)), CStr(rowArray(11)), 60
            enterRosterInfo CStr(rowArray(12)), CStr(rowArray(13)), 68
            enterRosterInfo CStr(rowArray(14)), CStr(rowArray(15)), 76
      
            wait
            Denver.screen.SendKeys ("<Enter>")
            wait

End Sub

Sub enterRosterInfo(st As String, et As String, co2 As Integer)
        
        Select Case st
        
            Case "SCK"
                Denver.screen.putstring "SIC ", 6, co2
                wait
                Denver.screen.putstring "    ", 7, co2
                wait
                Denver.screen.putstring "  ", 8, co2 + 2
                
            Case "HOL"
                Denver.screen.putstring "VAC ", 6, co2
                wait
                Denver.screen.putstring "    ", 7, co2
                wait
                Denver.screen.putstring "  ", 8, co2 + 2
                
            Case "OFF"
                Denver.screen.putstring "    ", 6, co2
                wait
                Denver.screen.putstring "    ", 7, co2
                wait
                Denver.screen.putstring "  ", 8, co2 + 2
                
            Case "UAB"
                Denver.screen.putstring "    ", 6, co2
                wait
                Denver.screen.putstring "    ", 7, co2
                wait
                Denver.screen.putstring "  ", 8, co2 + 2
                
            Case "ALM"
                Denver.screen.putstring "    ", 6, co2
                wait
                Denver.screen.putstring "    ", 7, co2
                wait
                Denver.screen.putstring "  ", 8, co2 + 2
                
            Case "AAB"
                Denver.screen.putstring "    ", 6, co2
                wait
                Denver.screen.putstring "    ", 7, co2
                wait
                Denver.screen.putstring "  ", 8, co2 + 2
                
            Case Else
            
                If checkIfTime(st) Then
                    Denver.screen.putstring st, 6, co2
                    wait
                    Denver.screen.putstring et, 7, co2
                    wait
                    Denver.screen.putstring getWorkedTime(st, et), 8, co2 + 2
                End If
                
        End Select
            
End Sub

Function checkIfTime(inOutTime As String) As Boolean
    
    checkIfTime = True
    Select Case inOutTime
    
        Case "ALM", "UAB", "HOL", "AAB", "SCK", "OFF", ""
            checkIfTime = False
        Case Else
            checkIfTime = True
    
    End Select
    
End Function

Function getWorkedTime(st As String, et As String) As String
    
    Dim result As Integer
    Dim startT As Integer
    Dim endT As Integer
    
    startT = CInt(st)
    endT = CInt(et)
    
    If endT > startT Then
        result = (endT - startT) / 100
    ElseIf endT < startT Then
        result = (endT + 2400 - startT) / 100
    End If
    
    If result < 10 Then
        getWorkedTime = "0" & CStr(result)
    Else
        getWorkedTime = "10"
    End If


End Function

Sub clearScreen_Click()

    Dim rowCount As Integer
    
    With ActiveSheet
        rowCount = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    
    ActiveSheet.Range("A2:P" & rowCount).Clear
    
End Sub
