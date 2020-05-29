Attribute VB_Name = "Module1"
Option Explicit
'Eddie Branigan 27/05/2020

Public g_HostSettleTime%
Public Sessions As Object
Public System As Object
Public Paperless As Object

Sub GetAssemblyReports_Click()

    Set System = CreateObject("BlueZone.System")
    Set Sessions = System.Sessions
    Set Paperless = Sessions.Item(1)
       
    If Paperless Is Nothing Then
    
        MsgBox "Couldn't connect to Paperless."
        
        Stop
    End If
    
    Paperless.screen.WaitHostQuiet (400)
    MsgBox (Paperless.screen.area(2, 33, 2, 48))
    Paperless.screen.WaitHostQuiet (400)
    If Paperless.screen.area(2, 33, 2, 48) = "7350 - Main Menu" Then
        navAssembly
        Paperless.screen.WaitHostQuiet (400)
        getAssemblyReport
    Else
        MsgBox ("Please return to Main Menu screen in Paperless before running macro.")
    End If
    
End Sub

Sub navAssembly()
       
    Paperless.screen.SendKeys ("ACCU<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    Paperless.screen.SendKeys ("ASSEMBLY<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    Paperless.screen.SendKeys ("Weekly<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    Paperless.screen.SendKeys ("<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    
End Sub

Sub getAssemblyReport()
    
    Dim firstDate As Date
    Dim currentDate As Date
    Dim checkLine As String
    Dim currentLine As Integer
    Dim workLine As Integer
    Dim x As Integer
    
    Paperless.screen.WaitHostQuiet (400)
    firstDate = CDate(Paperless.screen.area(12, 1, 12, 11))
    currentDate = firstDate
    workLine = 2
    
    For x = 13 To 23
            
        'problem if its not a date
        If IsDate(Paperless.screen.area(x, 1, x, 11)) Then
        
            currentDate = CDate(Paperless.screen.area(x, 1, x, 11))
            
        End If
        
        If regexTest(Paperless.screen.area(x, 1, x, 4)) Then
                
                setAssemblyLine Paperless.screen.area(x, 1, x, 4), _
                                Paperless.screen.area(x, 6, x, 15), _
                                Paperless.screen.area(x, 60, x, 67), _
                                Paperless.screen.area(x, 70, x, 77), _
                                CStr(currentDate), x, workLine
                workLine = workLine + 1
                
        End If
        
        If (x = 23) And Paperless.screen.area(x, 1, x, 5) = "=====" Then
            
            x = 3
            Paperless.screen.SendKeys ("N")
            Paperless.screen.WaitHostQuiet (400)
            
        End If
        
        If Paperless.screen.area(x, 1, x, 5) = "Note:" _
        And currentDate = firstDate + 6 Then
        
            x = 23
            
        End If
        
    Next x

End Sub

Function regexTest(screenSpace As String) As Boolean
    Dim regexOne As Object
    Set regexOne = New RegExp
     
    regexOne.Pattern = "\d{4}"
    
    regexTest = regexOne.test(screenSpace)
 
End Function

Sub setAssemblyLine(x1 As String, x2 As String, x3 As String, _
                    x4 As String, x5 As String, xL As Integer, _
                    workLine)
    
    ThisWorkbook.Worksheets("Assembly").Cells(workLine, 1) = x1
    ThisWorkbook.Worksheets("Assembly").Cells(workLine, 2) = x2
    ThisWorkbook.Worksheets("Assembly").Cells(workLine, 3) = x3
    ThisWorkbook.Worksheets("Assembly").Cells(workLine, 4) = x4
    ThisWorkbook.Worksheets("Assembly").Cells(workLine, 5) = x5

End Sub

Sub navGoodsIn()
       
    Paperless.screen.SendKeys ("ACCU<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    Paperless.screen.SendKeys ("Goods<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    Paperless.screen.SendKeys ("Weekly Goods-In Accu<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    Paperless.screen.SendKeys ("<ENTER>")
    Paperless.screen.WaitHostQuiet (400)
    
End Sub
