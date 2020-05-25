Attribute VB_Name = "Module1"
Option Explicit

Global objSessions          As Object
Global objSystem            As Object
Global Denver               As Object

Sub GetVolumes_Click()
    
    Set objSystem = CreateObject("BlueZone.System")
    Set objSessions = objSystem.Sessions
    Set Denver = objSystem.ActiveSession
    
    navDestination
    
    If Denver.screen.Area(1, 2, 1, 12) <> "CA View EXP" Then
    
        MsgBox ("Failed to reach TPX menu, please return to main screen")
        
    Else
        
        fillSARS01Screen
        getVolumeData
        exitReport
        
    End If
    
End Sub

Sub wait()

    Denver.screen.WaitHostQuiet (1)
    
End Sub

Sub navDenverMain()

    Denver.screen.SendKeys ("<RESET>")
    wait
    Denver.screen.SendKeys ("<PF3><PF3><PF3><PF3>")
    wait

End Sub

Sub navSARS01()

    Denver.screen.putString "SARI01", 23, 15
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait

End Sub

Sub fillSARS01Screen()

    Denver.screen.putString "JL0R06-OG", 6, 23
    Denver.screen.putString "ALL", 8, 48
    Denver.screen.putString "ALL", 14, 48
    Denver.screen.putString Format(Date - 1, "DDMMYY"), 20, 28
    Denver.screen.putString Format(Date, "DDMMYY"), 21, 28
    wait
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.putString "S", 6, 2
    Denver.screen.SendKeys ("<Enter>")
    wait
    Denver.screen.putString "find Ballymun"
    Denver.screen.SendKeys ("<Enter>")
    wait
    
End Sub

Sub getVolumeData()

    Dim shtRow As Integer
    Dim frozenLine As Integer
    Dim examineThis As String
    Dim sentinel As String
    Dim Line As Integer
    Dim wrksht As Worksheet
    
    Set wrksht = ThisWorkbook.Worksheets("Main")
    shtRow = 2
    examineThis = ""

    For Line = 3 To 26
        
        'check for group letter
        If Denver.screen.Area(Line, 1, Line, 11) = "ORDER GROUP" Then
        
            wrksht.Range("B" & shtRow) = Denver.screen.Area(Line, 16, Line, 16)
            
        End If
        
        'check for volume numbers
        If Denver.screen.Area(Line, 12, Line, 13) = "FU" Then
            
            wrksht.Range("A" & shtRow) = Denver.screen.Area(Line, 5, Line, 10)          'week
            wrksht.Range("C" & shtRow) = Denver.screen.Area(Line, 15, Line, 21)         'mon
            wrksht.Range("D" & shtRow) = Denver.screen.Area(Line, 23, Line, 29)         'tue
            wrksht.Range("E" & shtRow) = Denver.screen.Area(Line, 31, Line, 37)         'wed
            wrksht.Range("F" & shtRow) = Denver.screen.Area(Line, 39, Line, 45)         'thu
            wrksht.Range("G" & shtRow) = Denver.screen.Area(Line, 47, Line, 53)         'fri
            wrksht.Range("H" & shtRow) = Denver.screen.Area(Line, 55, Line, 61)         'sat
            wrksht.Range("I" & shtRow) = Denver.screen.Area(Line, 63, Line, 69)         'sun
            
            shtRow = shtRow + 1
            
        End If
        
        'go to next page
        If Line = 26 Then
        
            Line = 0
            Denver.screen.SendKeys ("<PF8>")
            wait
            
        End If
        
        
        If InStr(Denver.screen.Area(Line, 1, Line, 81), "BALLUMUN FROZEN") > 1 Then
            
            sentinel = "CR DEPOT DELIVERY VOLUME PREDICTION REPORT"
            Line = 26
            
            For frozenLine = 3 To 26
                
                'check for group letter
                If Denver.screen.Area(frozenLine, 1, frozenLine, 11) = "ORDER GROUP" Then
        
                    wrksht.Range("B" & shtRow) = Denver.screen.Area(frozenLine, 16, frozenLine, 16)
            
                End If
                
                'check for volume numbers
                If Denver.screen.Area(frozenLine, 12, frozenLine, 13) = "FU" Then
                    
                    wrksht.Range("A" & shtRow) = Denver.screen.Area(frozenLine, 5, frozenLine, 10)          'week
                    wrksht.Range("C" & shtRow) = Denver.screen.Area(frozenLine, 15, frozenLine, 21)         'mon
                    wrksht.Range("D" & shtRow) = Denver.screen.Area(frozenLine, 23, frozenLine, 29)         'tue
                    wrksht.Range("E" & shtRow) = Denver.screen.Area(frozenLine, 31, frozenLine, 37)         'wed
                    wrksht.Range("F" & shtRow) = Denver.screen.Area(frozenLine, 39, frozenLine, 45)         'thu
                    wrksht.Range("G" & shtRow) = Denver.screen.Area(frozenLine, 47, frozenLine, 53)         'fri
                    wrksht.Range("H" & shtRow) = Denver.screen.Area(frozenLine, 55, frozenLine, 61)         'sat
                    wrksht.Range("I" & shtRow) = Denver.screen.Area(frozenLine, 63, frozenLine, 69)         'sun
                    
                    shtRow = shtRow + 1
                    
                End If
                
                'go to next page
                If frozenLine = 26 Then
                
                    frozenLine = 0
                    Denver.screen.SendKeys ("<PF8>")
                    wait
                    
                End If
                
                If InStr(Denver.screen.Area(Line, 1, Line, 81), sentinel) > 1 Then
            
                    frozenLine = 26
                    
                End If
                
                Next frozenLine
            
        End If
    
    Next Line
    
End Sub

Sub exitReport()
    
    Denver.screen.SendKeys ("<RESET>")
    wait
    Denver.screen.SendKeys ("<PF3><PF3><PF3><PF3><PF3>")
    wait
    
End Sub

Sub navDestination()
    
    If Denver.screen.Area(1, 31, 1, 36) = "DENVER" Then
        
        Denver.screen.SendKeys ("<PA2>")
        wait
        navSARS01
        wait
        
    ElseIf Denver.screen.Area(1, 25, 1, 32) = "TPX MENU" Then
        
        navSARS01
        wait
        
    Else
        
        navDenverMain
        wait
        Denver.screen.SendKeys ("<PA2>")
        wait
        navSARS01
        wait
        
    End If

End Sub
