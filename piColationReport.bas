Attribute VB_Name = "Main"
Option Explicit

'Eddie Branigan 2020

Sub runMacro()

    Dim sentinel As Integer
    Dim weekNoRef As String
    Dim cell As Range
    Dim startRng As Range
    Dim currentRng As Range
    Dim currentFileName As String
    Dim filecount As Integer
    Dim dirFile As String
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    
    dirFile = "\\global.tesco.org\dfsroot\IE\Distribution\Ballymun\Planning\Warehouse Planning\PRA\PRA individual weeks\"
    
    clearWorkSheets
    sentinel = 1
    weekNoRef = ThisWorkbook.Worksheets("Directory").Cells(2, 2).Text

    Set startRng = ThisWorkbook.Worksheets("Directory").Range("B5:B265").Find(weekNoRef)
    Set currentRng = startRng

    Do Until sentinel = 26
        
        Set currentRng = currentRng.Offset(-1, 0)
        
        currentFileName = ""
        currentFileName = fileNameBuilder(currentRng.Text)
        
        If checkForFile(currentFileName, dirFile) Then
        
            fetchFile currentFileName, currentRng.Text, dirFile
            filecount = filecount + 1
            
        End If
        
        sentinel = sentinel + 1

    Loop
    
    ThisWorkbook.Worksheets("Summary").Cells(7, 7).Value = filecount
    
    'cleanData
    dcamDictionary (filecount)
    formatCols
     
End Sub

Function fileNameBuilder(cellRef As String) As String

    Dim yearString As String
    Dim weekString As String
    
    yearString = Left(cellRef, 4)
    weekString = Right(cellRef, 2)
    fileNameBuilder = yearString & " - WK " & weekString & ".xlsx"

End Function

Function checkForFile(fileName As String, dirFile As String) As Boolean
    
    Dim fileLoc As String
    fileLoc = dirFile & fileName
    
    If Len(Dir(fileLoc)) = 0 Then
        checkForFile = False
    Else
        checkForFile = True
    End If

End Function

Sub fetchFile(fileName As String, weekNumber As String, dirFile)

    Application.ScreenUpdating = False
    
    Dim praFile As Workbook
    
    Dim praRow As Integer
    Dim lastRow As Integer
    Dim praLastRow As Integer

    fileName = dirFile & fileName
    Set praFile = Workbooks.Open(fileName, UpdateLinks:=False, ReadOnly:=True)
    
    'pra file
    praLastRow = 0
    praLastRow = praFile.Worksheets("90-94.99 PI").Cells(Rows.Count, "A").End(xlUp).Row

    For praRow = 2 To praLastRow
    
            lastRow = ThisWorkbook.Worksheets("90-94.99").Cells(Rows.Count, "A").End(xlUp).Row + 1
        
            'WEEK NO
            ThisWorkbook.Worksheets("90-94.99").Range("A" & lastRow) = Right(weekNumber, 2)
            'DCAM
            ThisWorkbook.Worksheets("90-94.99").Range("B" & lastRow) = _
                praFile.Worksheets("90-94.99 PI").Range("A" & praRow).Text
            'NAME
            ThisWorkbook.Worksheets("90-94.99").Range("C" & lastRow) _
                = praFile.Worksheets("90-94.99 PI").Range("C" & praRow).Text & _
                " " & praFile.Worksheets("90-94.99 PI").Range("B" & praRow).Text
            'COMBINED PI
            ThisWorkbook.Worksheets("90-94.99").Range("D" & lastRow) = _
                praFile.Worksheets("90-94.99 PI").Range("E" & praRow).Text
            'CHECK
            ThisWorkbook.Worksheets("90-94.99").Range("E" & lastRow) = _
                praFile.Worksheets("90-94.99 PI").Range("I" & praRow).Text
    Next praRow
    
    
    'pra file
    praLastRow = 0
    praLastRow = praFile.Worksheets("95-97.99 PI").Cells(Rows.Count, "A").End(xlUp).Row
    
    For praRow = 2 To praLastRow
    
            lastRow = ThisWorkbook.Worksheets("95-97.99").Cells(Rows.Count, "A").End(xlUp).Row + 1
        
            'WEEK NO
            ThisWorkbook.Worksheets("95-97.99").Range("A" & lastRow) = Right(weekNumber, 2)
            'DCAM
            ThisWorkbook.Worksheets("95-97.99").Range("B" & lastRow) = _
                praFile.Worksheets("95-97.99 PI").Range("A" & praRow).Text
            'NAME
            ThisWorkbook.Worksheets("95-97.99").Range("C" & lastRow) _
                = praFile.Worksheets("95-97.99 PI").Range("C" & praRow).Text & _
                " " & praFile.Worksheets("95-97.99 PI").Range("B" & praRow).Text
            'COMBINED PI
            ThisWorkbook.Worksheets("95-97.99").Range("D" & lastRow) = _
                praFile.Worksheets("95-97.99 PI").Range("E" & praRow).Text
            'CHECK
            ThisWorkbook.Worksheets("95-97.99").Range("E" & lastRow) = _
                praFile.Worksheets("95-97.99 PI").Range("I" & praRow).Text
    Next praRow
    
    'pra file
    praLastRow = 0
    praLastRow = praFile.Worksheets(">=98 PI").Cells(Rows.Count, "A").End(xlUp).Row
    
    For praRow = 2 To praLastRow
    
            lastRow = ThisWorkbook.Worksheets("98+").Cells(Rows.Count, "A").End(xlUp).Row + 1
        
            'WEEK NO
            ThisWorkbook.Worksheets("98+").Range("A" & lastRow) = Right(weekNumber, 2)
            'DCAM
            ThisWorkbook.Worksheets("98+").Range("B" & lastRow) = _
                praFile.Worksheets(">=98 PI").Range("A" & praRow).Text
            'NAME
            ThisWorkbook.Worksheets("98+").Range("C" & lastRow) _
                = praFile.Worksheets(">=98 PI").Range("C" & praRow).Text & _
                " " & praFile.Worksheets(">=98 PI").Range("B" & praRow).Text
            'COMBINED PI
            ThisWorkbook.Worksheets("98+").Range("D" & lastRow) = _
                praFile.Worksheets(">=98 PI").Range("E" & praRow).Text
            'CHECK
            ThisWorkbook.Worksheets("98+").Range("E" & lastRow) = _
                praFile.Worksheets(">=98 PI").Range("I" & praRow).Text
    Next praRow
    
    praFile.Saved = True
    praFile.Close
    
    Application.ScreenUpdating = True

End Sub

Sub clearWorkSheets()
        
    ThisWorkbook.Worksheets("90-94.99").Range("A2:E1000").Clear
    ThisWorkbook.Worksheets("95-97.99").Range("A2:E1000").Clear
    ThisWorkbook.Worksheets("98+").Range("A2:E1000").Clear
    ThisWorkbook.Worksheets("Summary").Range("A3:A500").Clear
    ThisWorkbook.Worksheets("Summary").Range("C3:C500").Clear
    ThisWorkbook.Worksheets("Summary").Range("G7").Clear

End Sub

Sub cleanData()

    ThisWorkbook.Worksheets("90-94.99").Range("E:E") = [index(Upper(E:E),)]
    ThisWorkbook.Worksheets("95-97.99").Range("E:E") = [index(Upper(E:E),)]
    ThisWorkbook.Worksheets("98+").Range("E:E") = [index(Upper(E:E),)]
    
End Sub

Sub formatCols()

    ThisWorkbook.Worksheets("Summary").Range("C:C").NumberFormat = "0.00"

End Sub

Sub dcamDictionary(filecount As Integer)

    Dim dict As Scripting.Dictionary
    Dim ws90 As Worksheet
    Dim ws95 As Worksheet
    Dim ws98 As Worksheet
    Dim wsSum As Worksheet
    
    
    Dim x As Integer
    
    Set ws90 = ThisWorkbook.Worksheets("90-94.99")
    Set ws95 = ThisWorkbook.Worksheets("95-97.99")
    Set ws98 = ThisWorkbook.Worksheets("98+")
    Set wsSum = ThisWorkbook.Worksheets("Summary")
    Set dict = New Scripting.Dictionary

    dict.RemoveAll
  
    For x = 2 To ws90.Cells(Rows.Count, "A").End(xlUp).Row
        If ws90.Cells(x, 5).Text = "AGREED" Then
            If dict.Exists(ws90.Cells(x, 2).Text) Then
                dict(ws90.Cells(x, 2).Text) = ws90.Cells(x, 4).Value + dict(ws90.Cells(x, 2).Text)
            Else
                dict.Add ws90.Cells(x, 2).Text, ws90.Cells(x, 4).Value
            End If
        End If
    Next x
    
    For x = 2 To ws95.Cells(Rows.Count, "A").End(xlUp).Row
        If ws95.Cells(x, 5).Text = "AGREED" Then
            If dict.Exists(ws95.Cells(x, 2).Text) Then
                dict(ws95.Cells(x, 2).Text) = ws95.Cells(x, 4).Value + dict(ws95.Cells(x, 2).Text)
            Else
                dict.Add ws95.Cells(x, 2).Text, ws95.Cells(x, 4).Value
            End If
        End If
    Next x
    
    For x = 2 To ws98.Cells(Rows.Count, "A").End(xlUp).Row
        If ws98.Cells(x, 5).Text = "AGREED" Then
            If dict.Exists(ws98.Cells(x, 2).Text) Then
                dict(ws98.Cells(x, 2).Text) = ws98.Cells(x, 4).Value + dict(ws98.Cells(x, 2).Text)
            Else
                dict.Add ws98.Cells(x, 2).Text, ws98.Cells(x, 4).Value
            End If
        End If
    Next x
    
    Dim k As Variant
    Dim r As Integer
    r = 3
    
    For Each k In dict.Keys
        wsSum.Cells(r, 1) = k
        wsSum.Cells(r, 3) = dict(k)
        r = r + 1
    Next
    
    Dim praRange As Range
    Dim cell As Range
    
    Set praRange = wsSum.Range("C3:C" & wsSum.Cells(Rows.Count, "C").End(xlUp).Row)
    
    For Each cell In praRange
        cell.Value = cell.Value / filecount
        
    Next cell
    
    'Dim k As Variant
    'For Each k In dict.Keys
        ' Print key and value
    'Debug.Print k, dict(k)
    'Next
        
End Sub

Sub getNames()

    Dim wsSum As Worksheet
    wsSum = ThisWorkbook.Worksheets("Summary")
    Dim cell As Integer
    
    For cell = 3 To wsSum.Cells(Rows.Count, "C").End(xlUp).Row
        
        
    Next cell
    

End Sub
