Function Maxs(WorkRng As Range) As Double
'Update 20130907
Dim arr As Variant
arr = WorkRng.Value
For i = 1 To UBound(arr, 1)
    For j = 1 To UBound(arr, 2)
        arr(i, j) = VBA.Abs(arr(i, j))
    Next
Next
Maxs = Application.WorksheetFunction.Max(arr)
End Function
Function Mins(WorkRng As Range) As Double
Dim arr As Variant
arr = WorkRng.Value
For i = 1 To UBound(arr, 1)
    For j = 1 To UBound(arr, 2)
        arr(i, j) = VBA.Abs(arr(i, j))
    Next
Next
Mins = Application.WorksheetFunction.Min(arr)
End Function

Function Last(choice As Long, rng As Range)
' 1 = last row
' 2 = last column
' 3 = last cell
    Dim lrw As Long
    Dim lcol As Long

    Select Case choice

    Case 1:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
        On Error GoTo 0

    Case 2:
        On Error Resume Next
        Last = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

    Case 3:
        On Error Resume Next
        lrw = rng.Find(What:="*", _
                       After:=rng.Cells(1), _
                       Lookat:=xlPart, _
                       LookIn:=xlFormulas, _
                       SearchOrder:=xlByRows, _
                       SearchDirection:=xlPrevious, _
                       MatchCase:=False).Row
        On Error GoTo 0

        On Error Resume Next
        lcol = rng.Find(What:="*", _
                        After:=rng.Cells(1), _
                        Lookat:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
        On Error GoTo 0

        On Error Resume Next
        Last = rng.Parent.Cells(lrw, lcol).Address(False, False)
        If Err.Number > 0 Then
            Last = rng.Cells(1).Address(False, False)
            Err.Clear
        End If
        On Error GoTo 0

    End Select
End Function
Function IsTime(rng As Range) As Boolean
    Dim sValue As String
    sValue = rng.Cells(1).Text
    On Error Resume Next
    IsTime = IsDate(TimeValue(sValue))
    On Error GoTo 0
End Function

Sub Macro2()

Dim Data As Workbook
'Application.AskToUpdateLinks = False
Set Data = Workbooks.Open("C:\Users\surya.murali\Desktop\Macros\Data.xlsx")
Dim D As Worksheet
Set D = Data.Sheets("QA-348")

D.Range(D.Cells(3, 12), D.Cells(Last(1, D.Range("D3:D10000")), 12)).Select
Selection.NumberFormat = "[h]:mm:ss;@"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="Shafts"
Dim Results As Workbook
Application.AskToUpdateLinks = False
Set Results = Workbooks.Add
With Results
        .Title = "Result" & Date
        
    End With
Dim SHAFT As Worksheet
     Set SHAFT = Sheets.Add(After:=Sheets(Worksheets.Count))
     SHAFT.Name = "SHAFT"
Dim TULIP As Worksheet
     Set TULIP = Sheets.Add(After:=Sheets(Worksheets.Count))
     TULIP.Name = "TULIP"
Dim FiOR As Worksheet
     Set FiOR = Sheets.Add(After:=Sheets(Worksheets.Count))
     FiOR.Name = "FOR"
Dim INNERRACE As Worksheet
     Set INNERRACE = Sheets.Add(After:=Sheets(Worksheets.Count))
     INNERRACE.Name = "INNERRACE"
Dim CAGE As Worksheet
     Set CAGE = Sheets.Add(After:=Sheets(Worksheets.Count))
     CAGE.Name = "CAGE"
Dim SPIDER As Worksheet
     Set SPIDER = Sheets.Add(After:=Sheets(Worksheets.Count))
     SPIDER.Name = "SPIDER"
Dim BMW As Worksheet
     Set BMW = Sheets.Add(After:=Sheets(Worksheets.Count))
     BMW.Name = "BMW"
Dim Metric As Worksheet
     Set Metric = Sheets.Add(After:=Sheets(Worksheets.Count))
     Metric.Name = "Metric"
D.Activate
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="Shafts"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D1:D10000")), Last(2, D.Range("A2:EZ2")))).Copy
SHAFT.Activate
SHAFT.Cells(1, 1).PasteSpecial xlPasteValues
D.Activate
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="Tulips"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D1:D10000")), Last(2, D.Range("A2:EZ2")))).Copy
TULIP.Activate
TULIP.Cells(1, 1).Select
Selection.PasteSpecial
D.Activate
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="FOR"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D1:D10000")), Last(2, D.Range("A2:EZ2")))).Copy
FiOR.Activate
FiOR.Cells(1, 1).Select
Selection.PasteSpecial
D.Activate
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="FIR"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D1:D10000")), Last(2, D.Range("A2:EZ2")))).Copy
INNERRACE.Activate
INNERRACE.Cells(1, 1).Select
Selection.PasteSpecial
D.Activate
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="Cage"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D1:D10000")), Last(2, D.Range("A2:EZ2")))).Copy
CAGE.Activate
CAGE.Cells(1, 1).Select
Selection.PasteSpecial
D.Activate
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="SPIDER"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D1:D10000")), Last(2, D.Range("A2:EZ2")))).Copy
SPIDER.Activate
SPIDER.Cells(1, 1).Select
Selection.PasteSpecial
D.Activate
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D3:D10000")))).AutoFilter Field:=4, Criteria1:="BMW"
D.Range(D.Cells(2, 1), D.Cells(Last(1, D.Range("D1:D10000")), Last(2, D.Range("A2:EZ2")))).Copy
BMW.Activate
BMW.Cells(1, 1).Select
Selection.PasteSpecial
Metric.Activate
Metric.Cells(1, 1).Value = Date
Metric.Range("A1").Select
Selection.NumberFormat = "mmm-yy"
Metric.Cells(1, 2).Value = "V/S"
Metric.Cells(1, 3).Value = "Total Entries"
Metric.Cells(1, 4).Value = "Data Entry Errors"
Metric.Cells(1, 5).Value = "Rejects"
Metric.Cells(1, 6).Value = "Average Time For Rejected"
Metric.Cells(1, 8).Value = "Average Time"
Metric.Cells(1, 9).Value = "Max Time"
Metric.Cells(1, 10).Value = "Min Time"
Metric.Cells(1, 11).Value = "Missing or Corrupted Time"
Metric.Cells(1, 7).Value = "Multiple Rejects"
Metric.Cells(2, 2).Value = "Shaft"
Metric.Cells(3, 2).Value = "Tulip"
Metric.Cells(4, 2).Value = "FOR"
Metric.Cells(5, 2).Value = "FIR"
Metric.Cells(6, 2).Value = "Spider"
Metric.Cells(7, 2).Value = "Cages"
Metric.Cells(8, 2).Value = "BMW"
Metric.Cells(9, 2).Value = "Total"
Metric.Range("A1:K9").Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Metric.Range("B2:B8").Select
Range("B2:B8").Select
    Selection.Font.Bold = True
Range("A1:K1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
Metric.Cells(2, 3).Value = Last(1, SHAFT.Range("D1:D10000"))
SHAFT.Activate
SHAFT.Range(SHAFT.Cells(2, 12), SHAFT.Cells(Last(1, SHAFT.Range("K1:K10000")), 12)).Select
Selection.Copy
Range("O1").Select
Selection.PasteSpecial Paste:=xlPasteValues
Metric.Activate
Metric.Cells(2, 9).Value = Application.WorksheetFunction.Max(SHAFT.Range(SHAFT.Cells(2, 12), SHAFT.Cells(Last(1, SHAFT.Range("L1:L10000")), 12)))
Metric.Cells(2, 10).Value = Application.WorksheetFunction.Min(SHAFT.Range(SHAFT.Cells(2, 12), SHAFT.Cells(Last(1, SHAFT.Range("L1:L10000")), 12)))
Metric.Range("I2:J2").Select
Selection.NumberFormat = "[h]:mm:ss;@"
SHAFT.Columns("O").EntireColumn.Delete
Metric.Cells(3, 3).Value = Last(1, TULIP.Range("D1:D10000"))
TULIP.Activate
TULIP.Range(TULIP.Cells(2, 12), TULIP.Cells(Last(1, TULIP.Range("K1:K10000")), 12)).Select
Selection.Copy
Range("O1").Select
Selection.PasteSpecial Paste:=xlPasteValues
Metric.Activate
Metric.Cells(3, 9).Value = Application.WorksheetFunction.Max(TULIP.Range(TULIP.Cells(2, 12), TULIP.Cells(Last(1, TULIP.Range("L1:L10000")), 12)))
Metric.Cells(3, 10).Value = Application.WorksheetFunction.Min(TULIP.Range(TULIP.Cells(2, 12), TULIP.Cells(Last(1, TULIP.Range("L1:L10000")), 12)))
Metric.Range("I3:J3").Select
Selection.NumberFormat = "[h]:mm:ss;@"
TULIP.Columns("O").EntireColumn.Delete

Metric.Cells(4, 3).Value = Last(1, FiOR.Range("D1:D10000"))
FiOR.Activate
FiOR.Range(FiOR.Cells(2, 12), FiOR.Cells(Last(1, FiOR.Range("K1:K10000")), 12)).Select
Selection.Copy
Range("O1").Select
Selection.PasteSpecial Paste:=xlPasteValues
Metric.Activate
Metric.Cells(4, 9).Value = Application.WorksheetFunction.Max(FiOR.Range(FiOR.Cells(2, 12), FiOR.Cells(Last(1, FiOR.Range("L1:L10000")), 12)))
Metric.Cells(4, 10).Value = Application.WorksheetFunction.Min(FiOR.Range(FiOR.Cells(2, 12), FiOR.Cells(Last(1, FiOR.Range("L1:L10000")), 12)))
Metric.Range("I4:J4").Select
Selection.NumberFormat = "[h]:mm:ss;@"
FiOR.Columns("O").EntireColumn.Delete

Metric.Cells(5, 3).Value = Last(1, INNERRACE.Range("D1:D10000"))
INNERRACE.Activate
INNERRACE.Range(INNERRACE.Cells(2, 12), INNERRACE.Cells(Last(1, INNERRACE.Range("K1:K10000")), 12)).Select
Selection.Copy
Range("O1").Select
Selection.PasteSpecial Paste:=xlPasteValues
Metric.Activate
Metric.Cells(5, 9).Value = Application.WorksheetFunction.Max(INNERRACE.Range(INNERRACE.Cells(2, 12), INNERRACE.Cells(Last(1, INNERRACE.Range("L1:L10000")), 12)))
Metric.Cells(5, 10).Value = Application.WorksheetFunction.Min(INNERRACE.Range(INNERRACE.Cells(2, 12), INNERRACE.Cells(Last(1, INNERRACE.Range("L1:L10000")), 12)))
Metric.Range("I5:J5").Select
Selection.NumberFormat = "[h]:mm:ss;@"
INNERRACE.Columns("O").EntireColumn.Delete

Metric.Cells(6, 3).Value = Last(1, SPIDER.Range("D1:D10000"))
SPIDER.Activate
SPIDER.Range(SPIDER.Cells(2, 12), SPIDER.Cells(Last(1, SPIDER.Range("K1:K10000")), 12)).Select
Selection.Copy
Range("O1").Select
Selection.PasteSpecial Paste:=xlPasteValues
Metric.Activate
Metric.Cells(6, 9).Value = Application.WorksheetFunction.Max(SPIDER.Range(SPIDER.Cells(2, 12), SPIDER.Cells(Last(1, SPIDER.Range("L1:L10000")), 12)))
Metric.Cells(6, 10).Value = Application.WorksheetFunction.Min(SPIDER.Range(SPIDER.Cells(2, 12), SPIDER.Cells(Last(1, SPIDER.Range("L1:L10000")), 12)))
Metric.Range("I6:J6").Select
Selection.NumberFormat = "[h]:mm:ss;@"
SPIDER.Columns("O").EntireColumn.Delete

Metric.Cells(7, 3).Value = Last(1, CAGE.Range("D1:D10000"))
CAGE.Activate
CAGE.Range(CAGE.Cells(2, 12), CAGE.Cells(Last(1, CAGE.Range("K1:K10000")), 12)).Select
Selection.Copy
Range("O1").Select
Selection.PasteSpecial Paste:=xlPasteValues
Metric.Activate
Metric.Cells(7, 9).Value = Application.WorksheetFunction.Max(CAGE.Range(CAGE.Cells(2, 12), CAGE.Cells(Last(1, CAGE.Range("L1:L10000")), 12)))
Metric.Cells(7, 10).Value = Application.WorksheetFunction.Min(CAGE.Range(CAGE.Cells(2, 12), CAGE.Cells(Last(1, CAGE.Range("L1:L10000")), 12)))
Metric.Range("I7:J7").Select
Selection.NumberFormat = "[h]:mm:ss;@"
CAGE.Columns("O").EntireColumn.Delete

Metric.Cells(8, 3).Value = Last(1, BMW.Range("D1:D10000"))
BMW.Activate
BMW.Range(BMW.Cells(2, 12), BMW.Cells(Last(1, BMW.Range("K1:K10000")), 12)).Select
Selection.Copy
Range("O1").Select
Selection.PasteSpecial Paste:=xlPasteValues
Metric.Activate
Metric.Cells(8, 9).Value = Application.WorksheetFunction.Max(BMW.Range(BMW.Cells(2, 12), BMW.Cells(Last(1, BMW.Range("L1:L10000")), 12)))
Metric.Cells(8, 10).Value = Application.WorksheetFunction.Min(BMW.Range(BMW.Cells(2, 12), BMW.Cells(Last(1, BMW.Range("L1:L10000")), 12)))
Metric.Range("I8:J8").Select
Selection.NumberFormat = "[h]:mm:ss;@"
BMW.Columns("O").EntireColumn.Delete

Metric.Range("C9").Value = Application.Sum(Range(Cells(2, 3), Cells(8, 3)))
SHAFT.Activate
a = 0
For i = 2 To Last(1, SHAFT.Range("K1:K500"))
If SHAFT.Cells(i, 11).Value = "Reject" Then
    a = 1
    Exit For
End If
Next i
If a = 1 Then
SHAFT.Range(SHAFT.Cells(2, 1), SHAFT.Cells(Last(1, SHAFT.Range("K2:K10000")))).AutoFilter Field:=11, Criteria1:="Reject"

SHAFT.Range(SHAFT.Cells(2, 12), SHAFT.Cells(Last(1, SHAFT.Range("K1:K1000")), 12)).Select
Selection.Copy
    
    Range("O1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    SHAFT.Range("$A$1:$N$10000").AutoFilter Field:=11
    Sheets("Metric").Select
    Range("F2").Select
    Application.CutCopyMode = False
    Metric.Range("F2").Value = Application.WorksheetFunction.Average(SHAFT.Range(SHAFT.Cells(1, 15), SHAFT.Cells(Last(1, SHAFT.Range("O1:O1000")), 15)))
    Range("F2").Select
    Selection.NumberFormat = "[h]:mm:ss;@"
    Metric.Cells(2, 5).Value = Last(1, SHAFT.Range("O1:O10000"))
    SHAFT.Columns("O").EntireColumn.Delete
End If

TULIP.Activate
a = 0
For i = 2 To Last(1, TULIP.Range("K1:K500"))
If TULIP.Cells(i, 11).Value = "Reject" Then
    a = 1
    Exit For
End If
Next i
If a = 1 Then
TULIP.Range(TULIP.Cells(2, 1), TULIP.Cells(Last(1, TULIP.Range("K3:K10000")))).AutoFilter Field:=11, Criteria1:="Reject"

TULIP.Range(TULIP.Cells(2, 12), TULIP.Cells(Last(1, TULIP.Range("K1:K1000")), 12)).Select
Selection.Copy
    
    Range("O1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    TULIP.Range("$A$1:$N$10000").AutoFilter Field:=11
    Sheets("Metric").Select
    Range("F3").Select
    Application.CutCopyMode = False
    Metric.Range("F3").Value = Application.WorksheetFunction.Average(TULIP.Range(TULIP.Cells(1, 15), TULIP.Cells(Last(1, TULIP.Range("O1:O1000")), 15)))
    Range("F3").Select
    Selection.NumberFormat = "[h]:mm:ss;@"
    Metric.Cells(3, 5).Value = Last(1, TULIP.Range("O1:O10000"))
    TULIP.Columns("O").EntireColumn.Delete
End If
FiOR.Activate
a = 0
For i = 2 To Last(1, FiOR.Range("K1:K500"))
If FiOR.Cells(i, 11).Value = "Reject" Then
    a = 1
    Exit For
End If
Next i
If a = 1 Then
FiOR.Range(FiOR.Cells(2, 1), FiOR.Cells(Last(1, FiOR.Range("K3:K10000")))).AutoFilter Field:=11, Criteria1:="Reject"

FiOR.Range(FiOR.Cells(2, 12), FiOR.Cells(Last(1, FiOR.Range("K1:K1000")), 12)).Select
Selection.Copy
    
    Range("O1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    FiOR.Range("$A$1:$N$10000").AutoFilter Field:=11
    Sheets("Metric").Select
    Range("F4").Select
    Application.CutCopyMode = False
    Metric.Range("F4").Value = Application.WorksheetFunction.Average(FiOR.Range(FiOR.Cells(1, 15), FiOR.Cells(Last(1, FiOR.Range("O1:O1000")), 15)))
    Range("F4").Select
    Selection.NumberFormat = "[h]:mm:ss;@"
    Metric.Cells(4, 5).Value = Last(1, FiOR.Range("O1:O10000"))
    FiOR.Columns("O").EntireColumn.Delete
End If
INNERRACE.Activate
a = 0
For i = 2 To Last(1, INNERRACE.Range("K1:K500"))
If INNERRACE.Cells(i, 11).Value = "Reject" Then
    a = 1
    Exit For
End If
Next i
If a = 1 Then
INNERRACE.Range(INNERRACE.Cells(2, 1), INNERRACE.Cells(Last(1, INNERRACE.Range("K3:K10000")))).AutoFilter Field:=11, Criteria1:="Reject"

INNERRACE.Range(INNERRACE.Cells(2, 12), INNERRACE.Cells(Last(1, INNERRACE.Range("K1:K1000")), 12)).Select
Selection.Copy
    
    Range("O1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    INNERRACE.Range("$A$1:$N$10000").AutoFilter Field:=11
    Sheets("Metric").Select
    Range("F5").Select
    Application.CutCopyMode = False
    Metric.Range("F5").Value = Application.WorksheetFunction.Average(INNERRACE.Range(INNERRACE.Cells(1, 15), INNERRACE.Cells(Last(1, INNERRACE.Range("O1:O1000")), 15)))
    Range("F5").Select
    Selection.NumberFormat = "[h]:mm:ss;@"
    Metric.Cells(5, 5).Value = Last(1, INNERRACE.Range("O1:O10000"))
    INNERRACE.Columns("O").EntireColumn.Delete
End If
SPIDER.Activate
a = 0
For i = 2 To Last(1, SPIDER.Range("K1:K500"))
If SPIDER.Cells(i, 11).Value = "Reject" Then
    a = 1
    Exit For
End If
Next i
If a = 1 Then
SPIDER.Range(SPIDER.Cells(2, 1), SPIDER.Cells(Last(1, SHAFT.Range("K3:K10000")))).AutoFilter Field:=11, Criteria1:="Reject"

SPIDER.Range(SPIDER.Cells(2, 12), SPIDER.Cells(Last(1, SPIDER.Range("K1:K1000")), 12)).Select
Selection.Copy
    
    Range("O1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    SPIDER.Range("$A$1:$N$10000").AutoFilter Field:=11
    Sheets("Metric").Select
    Range("F6").Select
    Application.CutCopyMode = False
    Metric.Range("F6").Value = Application.WorksheetFunction.Average(SPIDER.Range(SPIDER.Cells(1, 15), SPIDER.Cells(Last(1, SPIDER.Range("O1:O1000")), 15)))
    Range("F6").Select
    Selection.NumberFormat = "[h]:mm:ss;@"
    Metric.Cells(6, 5).Value = Last(1, SPIDER.Range("O1:O10000"))
    SPIDER.Columns("O").EntireColumn.Delete
End If
CAGE.Activate
a = 0
For i = 2 To Last(1, CAGE.Range("K1:K500"))
If CAGE.Cells(i, 11).Value = "Reject" Then
    a = 1
    Exit For
End If
Next i
If a = 1 Then
CAGE.Range(CAGE.Cells(2, 1), CAGE.Cells(Last(1, CAGE.Range("K3:K10000")))).AutoFilter Field:=11, Criteria1:="Reject"

CAGE.Range(CAGE.Cells(2, 12), CAGE.Cells(Last(1, CAGE.Range("K1:K1000")), 12)).Select
Selection.Copy
    
    Range("O1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    CAGE.Range("$A$1:$N$10000").AutoFilter Field:=11
    Sheets("Metric").Select
    Range("F7").Select
    Application.CutCopyMode = False
    Metric.Range("F7").Value = Application.WorksheetFunction.Average(CAGE.Range(CAGE.Cells(1, 15), CAGE.Cells(Last(1, CAGE.Range("O1:O1000")), 15)))
    Range("F7").Select
    Selection.NumberFormat = "[h]:mm:ss;@"
    Metric.Cells(7, 5).Value = Last(1, CAGE.Range("O1:O10000"))
    CAGE.Columns("O").EntireColumn.Delete
End If
BMW.Activate
a = 0
For i = 2 To Last(1, BMW.Range("K1:K500"))
If BMW.Cells(i, 11).Value = "Reject" Then
    a = 1
    Exit For
End If
Next i
If a = 1 Then
    BMW.Range(BMW.Cells(2, 1), BMW.Cells(Last(1, BMW.Range("K3:K10000")))).AutoFilter Field:=11, Criteria1:="Reject"

    BMW.Range(BMW.Cells(2, 12), BMW.Cells(Last(1, BMW.Range("K1:K1000")), 12)).Select
    Selection.Copy
        
        Range("O1").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        BMW.Range("$A$1:$N$10000").AutoFilter Field:=11
        Sheets("Metric").Select
        Range("F8").Select
        Application.CutCopyMode = False
        Metric.Range("F8").Value = Application.WorksheetFunction.Average(BMW.Range(BMW.Cells(1, 15), BMW.Cells(Last(1, BMW.Range("O1:O1000")), 15)))
        Range("F8").Select
        Selection.NumberFormat = "[h]:mm:ss;@"
        Metric.Cells(8, 5).Value = Last(1, BMW.Range("O1:O10000"))
        BMW.Columns("O").EntireColumn.Delete
End If
Metric.Range("E9").Value = Application.Sum(Metric.Range(Metric.Cells(2, 5), Metric.Cells(8, 5)))

Metric.Activate

Metric.Range("G:H,K:K").EntireColumn.Hidden = True


End Sub

