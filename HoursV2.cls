VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub setup()
    copyAndRename
    clearBody
    setDates
    unMergeTotals
End Sub

Sub copyAndRename()
    d = Date
    m = Month(d)
    WorksheetName = MonthName(m) & " " & Year(d)
    ActiveSheet.Copy Before:=Worksheets(1)
    On Error Resume Next
    ActiveSheet.Name = WorksheetName
End Sub

Sub clearBody()
    Dim rng As Range, cell As Range, clearRange As Range
    Set rng = Range("A1:A1000")
    
    For Each cell In rng
        If cell.Value = "Total" Then
            Set clearRange = Range("B1:AF" & cell.Row - 1)
            clearRange.ClearContents
            clearRange.ClearNotes
            Exit For
        End If
    Next cell
End Sub

Sub setDates()
    Dim rng As Range, i As Integer
            
    theDate = Date
    theMonth = Month(theDate)
    theYear = Year(theDate)
    
    daysInMonth = MonthDays(CLng(theMonth))
        
    For i = 1 To daysInMonth
        Cells(1, i + 1).Value = theMonth & "/" & i & "/" & theYear
    Next i
End Sub

Sub unMergeTotals()
    Dim rng As Range, cell As Range, mergeRow As Integer, mergeRange As Range

    theDate = Date
    theMonth = Month(theDate)
    theYear = Year(theDate)
    
    daysInMonth = MonthDays(CLng(theMonth))

    Set rng = Range("A1:A1000")
    
    For Each cell In rng
        If cell.Value = "Week Totals" Then
            mergeRow = cell.Row
            Exit For
        End If
    Next cell
    
    Set mergeRange = Range("B" & mergeRow & ":AF" & mergeRow)
    mergeRange.UnMerge
End Sub

Function MonthDays(myMonth As Long) As Long
   
    MonthDays = Day(DateSerial(Year(Date), myMonth + 1, 1) - 1)
     
End Function
