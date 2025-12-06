' https://chatgpt.com/c/6932279d-0234-8321-b5e4-f5943cb3cd8e

Option Explicit

Const START_DATE_CELL As String = "B2"
Const WINDOW_DAYS As Long = 3              ' ± days around the target date
Const DATE_FMT As String = "yyyy-mm-dd"    ' ISO format, e.g. 2025-12-05

' Adjust these storage conditions to match your climatic zone and protocol
Const LONG_TERM_COND As String = "Long-term (e.g. 25°C / 60% RH or 30°C / 75% RH)"
Const INTERMEDIATE_COND As String = "Intermediate (e.g. 30°C / 65% RH)"
Const ACCELERATED_COND As String = "Accelerated (e.g. 40°C / 75% RH)"

' === MAIN MACRO: build ICH-style schedule (Long-term, Accelerated, Intermediate) ===
Sub CreateICHStabilitySchedule()
    Dim ws As Worksheet
    Dim startDate As Date
    Dim longTermMonths As Variant
    Dim acceleratedMonths As Variant
    Dim intermediateMonths As Variant
    Dim nextRow As Long
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    
    ' Read start date
    If Not IsDate(ws.Range(START_DATE_CELL).Value) Then
        MsgBox "Please enter a valid START DATE in cell " & START_DATE_CELL & _
               " (for example 2025-12-05).", vbExclamation
        Exit Sub
    End If
    startDate = CDate(ws.Range(START_DATE_CELL).Value)
    
    ' Time points according to ICH-style schedules:
    ' Long-term:      0, 3, 6, 9, 12, 18, 24, 36 months
    ' Accelerated:    0, 3, 6 months
    ' Intermediate:   0, 6, 9, 12 months
    longTermMonths = Array(0, 3, 6, 9, 12, 18, 24, 36)
    acceleratedMonths = Array(0, 3, 6)
    intermediateMonths = Array(0, 6, 9, 12)
    
    ' Optional: clear previous table (A1:K500)
    ws.Range("A1:K500").ClearContents
    ws.Range("A1:K500").Interior.ColorIndex = xlNone
    
    ' Header row
    ws.Range("A1").Value = "Study type"
    ws.Range("B1").Value = "Storage condition"
    ws.Range("C1").Value = "Nominal time point"
    ws.Range("D1").Value = "Offset (months)"
    ws.Range("E1").Value = "Target date"
    ws.Range("F1").Value = "Earliest date (-" & WINDOW_DAYS & " days)"
    ws.Range("G1").Value = "Latest date (+" & WINDOW_DAYS & " days)"
    ws.Range("H1").Value = "Recommended pull date (Mon–Fri)"
    ws.Range("I1").Value = "Actual pull date"
    ws.Range("J1").Value = "Analyst initials"
    ws.Range("K1").Value = "Remarks"
    
    nextRow = 2
    
    ' Long-term block
    nextRow = AddScheduleBlock(ws, "LONG-TERM", LONG_TERM_COND, longTermMonths, startDate, nextRow)
    
    ' Blank row between blocks
    nextRow = nextRow + 1
    
    ' Accelerated block
    nextRow = AddScheduleBlock(ws, "ACCELERATED", ACCELERATED_COND, acceleratedMonths, startDate, nextRow)
    
    ' Blank row between blocks
    nextRow = nextRow + 1
    
    ' Intermediate block
    nextRow = AddScheduleBlock(ws, "INTERMEDIATE", INTERMEDIATE_COND, intermediateMonths, startDate, nextRow)
    
    ' Format dates and adjust column width
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow >= 2 Then
        ws.Range("E2:I" & lastRow).NumberFormat = DATE_FMT
        ws.Columns("A:K").AutoFit
    End If
    
    MsgBox "ICH-style stability schedule created from start date " & _
           Format(startDate, DATE_FMT) & ".", vbInformation
End Sub

' === Helper: add one block (Long-term / Accelerated / Intermediate) ===
Private Function AddScheduleBlock( _
    ByVal ws As Worksheet, _
    ByVal studyType As String, _
    ByVal storageCond As String, _
    ByVal monthsArray As Variant, _
    ByVal startDate As Date, _
    ByVal startRow As Long) As Long
    
    Dim i As Long
    Dim r As Long
    Dim m As Long
    Dim label As String
    Dim targetDate As Date
    Dim earliestDate As Date
    Dim latestDate As Date
    Dim recDate As Date
    
    r = startRow
    
    For i = LBound(monthsArray) To UBound(monthsArray)
        m = CLng(monthsArray(i))
        label = DescribeTimePoint(m)
        targetDate = DateAdd("m", m, startDate)
        earliestDate = targetDate - WINDOW_DAYS
        latestDate = targetDate + WINDOW_DAYS
        recDate = RecommendedPullDate(targetDate, earliestDate, latestDate)
        
        ws.Cells(r, "A").Value = studyType
        ws.Cells(r, "B").Value = storageCond
        ws.Cells(r, "C").Value = label
        ws.Cells(r, "D").Value = m
        ws.Cells(r, "E").Value = targetDate
        ws.Cells(r, "F").Value = earliestDate
        ws.Cells(r, "G").Value = latestDate
        ws.Cells(r, "H").Value = recDate
        
        r = r + 1
    Next i
    
    AddScheduleBlock = r
End Function

' === Helper: make a nice label such as "12 months (1 year)" ===
Private Function DescribeTimePoint(ByVal months As Long) As String
    Dim result As String
    If months = 0 Then
        result = "0 month (initial)"
    ElseIf months < 12 Then
        result = CStr(months) & " months"
    ElseIf months Mod 12 = 0 Then
        result = CStr(months) & " months (" & CStr(months \ 12) & " year(s))"
    Else
        result = CStr(months) & " months (" & Format(months / 12, "0.0") & " years)"
    End If
    DescribeTimePoint = result
End Function

' === Workday helpers ===
Private Function IsWorkday(ByVal d As Date) As Boolean
    ' Weekday with Monday as 1, Sunday as 7
    Dim wd As Long
    wd = Weekday(d, vbMonday)
    IsWorkday = (wd >= 1 And wd <= 5)   ' Mon–Fri
End Function

Private Function RecommendedPullDate( _
    ByVal targetDate As Date, _
    ByVal earliestDate As Date, _
    ByVal latestDate As Date) As Date
    
    Dim d As Date
    
    ' If target is already a workday within the window, use it
    If IsWorkday(targetDate) And targetDate >= earliestDate And targetDate <= latestDate Then
        RecommendedPullDate = targetDate
        Exit Function
    End If
    
    ' First try searching backwards from target
    d = targetDate
    Do While d >= earliestDate
        d = d - 1
        If IsWorkday(d) Then
            RecommendedPullDate = d
            Exit Function
        End If
    Loop
    
    ' If nothing backwards, search forwards from target
    d = targetDate
    Do While d <= latestDate
        d = d + 1
        If IsWorkday(d) Then
            RecommendedPullDate = d
            Exit Function
        End If
    Loop
    
    ' Fallback (should rarely be used): return target date
    RecommendedPullDate = targetDate
End Function

' === CHECK MACRO: highlight pulls due today within ±WINDOW_DAYS ===
Sub CheckDueSamples_ICH()
    Dim ws As Worksheet
    Dim today As Date
    Dim lastRow As Long
    Dim i As Long
    Dim msg As String
    Dim hasHit As Boolean
    
    Set ws = ActiveSheet
    today = Date
    
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No schedule data found in column E.", vbExclamation
        Exit Sub
    End If
    
    ' Clear previous highlighting
    ws.Range("A2:K" & lastRow).Interior.ColorIndex = xlNone
    
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, "F").Value) And IsDate(ws.Cells(i, "G").Value) Then
            If today >= ws.Cells(i, "F").Value And today <= ws.Cells(i, "G").Value Then
                ws.Range("A" & i & ":K" & i).Interior.Color = vbYellow
                hasHit = True
                msg = msg & vbCrLf & _
                      ws.Cells(i, "A").Value & " | " & _
                      ws.Cells(i, "C").Value & " | target: " & _
                      Format(ws.Cells(i, "E").Value, DATE_FMT) & _
                      " | recommended: " & _
                      Format(ws.Cells(i, "H").Value, DATE_FMT)
            End If
        End If
    Next i
    
    If hasHit Then
        MsgBox "TODAY (" & Format(today, DATE_FMT) & _
               ") is within ±" & WINDOW_DAYS & _
               " days for the following pulls:" & vbCrLf & msg, vbInformation
    Else
        MsgBox "Today is not within ±" & WINDOW_DAYS & _
               " days of any planned pull.", vbInformation
    End If
End Sub
