
' https://chatgpt.com/c/6933935b-8db4-8320-b2de-49c3da8260a5

Option Explicit

' ==========================
' Global settings
' ==========================
Const START_DATE_CELL As String = "B2"
Const WINDOW_DAYS As Long = 3          ' ± days around target date

Const LANG_EN As String = "EN"
Const LANG_VI As String = "VI"

' ==========================
' PUBLIC ENTRY MACROS – BUILD SCHEDULE
' ==========================

' English, GMP-style dd-mmm-yyyy (e.g. 05-Dec-2025)
Public Sub ICH_Schedule_EN_dd_mmm()
    BuildICHStabilitySchedule LANG_EN, "dd-mmm-yyyy"
End Sub

' English, ISO format yyyy-mm-dd (e.g. 2025-12-05)
Public Sub ICH_Schedule_EN_ISO()
    BuildICHStabilitySchedule LANG_EN, "yyyy-mm-dd"
End Sub

' Vietnamese (khong dau), dd/mm/yyyy (e.g. 05/12/2025)
Public Sub ICH_Schedule_VI_dd_mm()
    BuildICHStabilitySchedule LANG_VI, "dd/mm/yyyy"
End Sub

' Vietnamese (khong dau), ISO format yyyy-mm-dd
Public Sub ICH_Schedule_VI_ISO()
    BuildICHStabilitySchedule LANG_VI, "yyyy-mm-dd"
End Sub

' ==========================
' PUBLIC ENTRY MACROS – CHECK DUE SAMPLES
' ==========================

Public Sub ICH_CheckDueSamples_EN()
    CheckDueSamples_ICH_Generic LANG_EN
End Sub

Public Sub ICH_CheckDueSamples_VI()
    CheckDueSamples_ICH_Generic LANG_VI
End Sub

' ==========================
' CORE BUILDER
' ==========================

Private Sub BuildICHStabilitySchedule( _
    ByVal lang As String, _
    ByVal dateFmt As String)

    Dim ws As Worksheet
    Dim startDate As Date
    Dim longTermMonths As Variant
    Dim acceleratedMonths As Variant
    Dim intermediateMonths As Variant
    Dim nextRow As Long
    Dim lastRow As Long
    
    Dim longTermStudy As String
    Dim intermediateStudy As String
    Dim acceleratedStudy As String
    
    Dim longTermCond As String
    Dim intermediateCond As String
    Dim acceleratedCond As String
    
    Set ws = ActiveSheet
    
    ' --- Read start date ---
    If Not IsDate(ws.Range(START_DATE_CELL).Value) Then
        If lang = LANG_VI Then
            MsgBox "Vui long nhap NGAY BAT DAU hop le o o " & START_DATE_CELL & _
                   " (vi du " & Format(DateSerial(2025, 12, 5), dateFmt) & ").", _
                   vbExclamation
        Else
            MsgBox "Please enter a valid START DATE in cell " & START_DATE_CELL & _
                   " (for example " & Format(DateSerial(2025, 12, 5), dateFmt) & ").", _
                   vbExclamation
        End If
        Exit Sub
    End If
    startDate = CDate(ws.Range(START_DATE_CELL).Value)
    
    ' --- ICH-style time points ---
    ' Long-term:      0, 3, 6, 9, 12, 18, 24, 36 months
    ' Accelerated:    0, 3, 6 months
    ' Intermediate:   0, 6, 9, 12 months
    longTermMonths = Array(0, 3, 6, 9, 12, 18, 24, 36)
    acceleratedMonths = Array(0, 3, 6)
    intermediateMonths = Array(0, 6, 9, 12)
    
    ' --- Define labels by language ---
    If lang = LANG_VI Then
        ' Study type labels – Vietnamese (khong dau, chu dau dong viet hoa)
        longTermStudy = "On Dinh Dai Han"
        intermediateStudy = "On Dinh Trung Gian"
        acceleratedStudy = "On Dinh Tang Toc"
        
        ' Storage conditions – Vietnamese (khong dau)
        longTermCond = "Dai Han (vi du 30°C / 75% RH)"
        intermediateCond = "Trung Gian (vi du 30°C / 65% RH)"
        acceleratedCond = "Tang Toc (vi du 40°C / 75% RH)"
    Else
        ' Study type labels – English
        longTermStudy = "Long-term Stability"
        intermediateStudy = "Intermediate Stability"
        acceleratedStudy = "Accelerated Stability"
        
        ' Storage conditions – English
        longTermCond = "Long-term (e.g. 25°C / 60% RH or 30°C / 75% RH)"
        intermediateCond = "Intermediate (e.g. 30°C / 65% RH)"
        acceleratedCond = "Accelerated (e.g. 40°C / 75% RH)"
    End If
    
    ' --- Clear previous table (A1:K500) ---
    ws.Range("A1:K500").ClearContents
    ws.Range("A1:K500").Interior.ColorIndex = xlNone
    
    ' --- Header row ---
    If lang = LANG_VI Then
        ' Tieu de cot: Tieng Viet khong dau, moi tu viet hoa chu dau
        ws.Range("A1").Value = "Loai Nghien Cuu"
        ws.Range("B1").Value = "Dieu Kien Bao Quan"
        ws.Range("C1").Value = "Moc Danh Nghia"
        ws.Range("D1").Value = "So Thang Bu"
        ws.Range("E1").Value = "Ngay Muc Tieu"
        ws.Range("F1").Value = "Ngay Som Nhat (-" & WINDOW_DAYS & " Ngay)"
        ws.Range("G1").Value = "Ngay Tre Nhat (+" & WINDOW_DAYS & " Ngay)"
        ws.Range("H1").Value = "Ngay Lay Mau De Xuat (T2–T6)"
        ws.Range("I1").Value = "Ngay Lay Mau Thuc Te"
        ws.Range("J1").Value = "Ky Hieu Kiem Nghiem Vien"
        ws.Range("K1").Value = "Ghi Chu"
    Else
        ws.Range("A1").Value = "Study Type"
        ws.Range("B1").Value = "Storage Condition"
        ws.Range("C1").Value = "Nominal Time Point"
        ws.Range("D1").Value = "Offset (Months)"
        ws.Range("E1").Value = "Target Date"
        ws.Range("F1").Value = "Earliest Date (-" & WINDOW_DAYS & " Days)"
        ws.Range("G1").Value = "Latest Date (+" & WINDOW_DAYS & " Days)"
        ws.Range("H1").Value = "Recommended Pull Date (Mon–Fri)"
        ws.Range("I1").Value = "Actual Pull Date"
        ws.Range("J1").Value = "Analyst Initials"
        ws.Range("K1").Value = "Remarks"
    End If
    
    nextRow = 2
    
    ' --- Long-term block ---
    nextRow = AddScheduleBlock(ws, longTermStudy, longTermCond, _
                               longTermMonths, startDate, nextRow, lang)
    
    ' Blank row between blocks
    nextRow = nextRow + 1
    
    ' --- Accelerated block ---
    nextRow = AddScheduleBlock(ws, acceleratedStudy, acceleratedCond, _
                               acceleratedMonths, startDate, nextRow, lang)
    
    ' Blank row between blocks
    nextRow = nextRow + 1
    
    ' --- Intermediate block ---
    nextRow = AddScheduleBlock(ws, intermediateStudy, intermediateCond, _
                               intermediateMonths, startDate, nextRow, lang)
    
    ' --- Format dates and autofit ---
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow >= 2 Then
        ws.Range("E2:I" & lastRow).NumberFormat = dateFmt
        ws.Columns("A:K").AutoFit
    End If
    
    ' --- Finished message ---
    If lang = LANG_VI Then
        MsgBox "Da tao lich do on dinh ICH tu ngay bat dau " & _
               Format(startDate, dateFmt) & ".", vbInformation
    Else
        MsgBox "ICH-style stability schedule created from start date " & _
               Format(startDate, dateFmt) & ".", vbInformation
    End If
End Sub

' ==========================
' Add one block (Long-term / Accelerated / Intermediate)
' ==========================

Private Function AddScheduleBlock( _
    ByVal ws As Worksheet, _
    ByVal studyType As String, _
    ByVal storageCond As String, _
    ByVal monthsArray As Variant, _
    ByVal startDate As Date, _
    ByVal startRow As Long, _
    ByVal lang As String) As Long
    
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
        label = DescribeTimePointGeneric(m, lang)
        
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

' ==========================
' Time-point description (EN / VI khong dau)
' ==========================

Private Function DescribeTimePointGeneric( _
    ByVal months As Long, _
    ByVal lang As String) As String
    
    Dim result As String
    
    If lang = LANG_VI Then
        If months = 0 Then
            result = "0 Thang (Ban Dau)"
        ElseIf months < 12 Then
            result = CStr(months) & " Thang"
        ElseIf months Mod 12 = 0 Then
            result = CStr(months) & " Thang (" & CStr(months \ 12) & " Nam)"
        Else
            result = CStr(months) & " Thang (" & Format(months / 12, "0.0") & " Nam)"
        End If
    Else
        If months = 0 Then
            result = "0 Month (Initial)"
        ElseIf months < 12 Then
            result = CStr(months) & " Months"
        ElseIf months Mod 12 = 0 Then
            result = CStr(months) & " Months (" & CStr(months \ 12) & " Year(s))"
        Else
            result = CStr(months) & " Months (" & Format(months / 12, "0.0") & " Years)"
        End If
    End If
    
    DescribeTimePointGeneric = result
End Function

' ==========================
' Workday helper – Mon–Fri only
' ==========================

Private Function IsWorkday(ByVal d As Date) As Boolean
    ' Weekday with Monday as 1, Sunday as 7
    Dim wd As Long
    wd = Weekday(d, vbMonday)
    
    ' Company works from Monday to Friday
    IsWorkday = (wd >= 1 And wd <= 5)    ' Mon–Fri
    
    ' If you want to allow Saturday as a working day:
    ' IsWorkday = (wd >= 1 And wd <= 6)
End Function

' ==========================
' Recommended pull date inside ±WINDOW_DAYS, Mon–Fri only
' ==========================

Private Function RecommendedPullDate( _
    ByVal targetDate As Date, _
    ByVal earliestDate As Date, _
    ByVal latestDate As Date) As Date
    
    Dim d As Date
    
    ' If target is already a workday within the window, use it
    If IsWorkday(targetDate) _
       And targetDate >= earliestDate _
       And targetDate <= latestDate Then
        RecommendedPullDate = targetDate
        Exit Function
    End If
    
    ' Search backwards from target-1 down to earliest
    d = targetDate - 1
    Do While d >= earliestDate
        If IsWorkday(d) Then
            RecommendedPullDate = d
            Exit Function
        End If
        d = d - 1
    Loop
    
    ' If nothing backwards, search forwards from target+1 up to latest
    d = targetDate + 1
    Do While d <= latestDate
        If IsWorkday(d) Then
            RecommendedPullDate = d
            Exit Function
        End If
        d = d + 1
    Loop
    
    ' Fallback (rare) – just return the target date
    RecommendedPullDate = targetDate
End Function

' ==========================
' Check which pulls are due today (± WINDOW_DAYS)
' ==========================

Private Sub CheckDueSamples_ICH_Generic(ByVal lang As String)
    Dim ws As Worksheet
    Dim today As Date
    Dim lastRow As Long
    Dim i As Long
    Dim msg As String
    Dim hasHit As Boolean
    Dim dateFmt As String
    
    Set ws = ActiveSheet
    today = Date
    
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    If lastRow < 2 Then
        If lang = LANG_VI Then
            MsgBox "Chua co du lieu lich on dinh trong cot E.", vbExclamation
        Else
            MsgBox "No schedule data found in column E.", vbExclamation
        End If
        Exit Sub
    End If
    
    ' Try to read date format from E2 (where schedule dates start)
    dateFmt = ws.Range("E2").NumberFormat
    If dateFmt = "" Or dateFmt = "General" Then
        dateFmt = "yyyy-mm-dd"
    End If
    
    ' Clear previous highlighting
    ws.Range("A2:K" & lastRow).Interior.ColorIndex = xlNone
    
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, "F").Value) And IsDate(ws.Cells(i, "G").Value) Then
            If today >= ws.Cells(i, "F").Value And today <= ws.Cells(i, "G").Value Then
                ws.Range("A" & i & ":K" & i).Interior.Color = vbYellow
                hasHit = True
                
                If lang = LANG_VI Then
                    msg = msg & vbCrLf & _
                          ws.Cells(i, "A").Value & " | " & _
                          ws.Cells(i, "C").Value & " | Muc Tieu: " & _
                          Format(ws.Cells(i, "E").Value, dateFmt) & _
                          " | De Xuat: " & _
                          Format(ws.Cells(i, "H").Value, dateFmt)
                Else
                    msg = msg & vbCrLf & _
                          ws.Cells(i, "A").Value & " | " & _
                          ws.Cells(i, "C").Value & " | Target: " & _
                          Format(ws.Cells(i, "E").Value, dateFmt) & _
                          " | Recommended: " & _
                          Format(ws.Cells(i, "H").Value, dateFmt)
                End If
            End If
        End If
    Next i
    
    If hasHit Then
        If lang = LANG_VI Then
            MsgBox "Hom nay (" & Format(today, dateFmt) & _
                   ") nam trong ±" & WINDOW_DAYS & _
                   " ngay cua cac moc sau:" & vbCrLf & msg, vbInformation
        Else
            MsgBox "TODAY (" & Format(today, dateFmt) & _
                   ") is within ±" & WINDOW_DAYS & _
                   " days for the following pulls:" & vbCrLf & msg, vbInformation
        End If
    Else
        If lang = LANG_VI Then
            MsgBox "Hom nay khong nam trong ±" & WINDOW_DAYS & _
                   " ngay cua moc nao.", vbInformation
        Else
            MsgBox "Today is not within ±" & WINDOW_DAYS & _
                   " days of any planned pull.", vbInformation
        End If
    End If
End Sub
