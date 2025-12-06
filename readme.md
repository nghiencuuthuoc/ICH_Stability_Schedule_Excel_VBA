
# Creating an ICH Stability Sampling Schedule in Excel with VBA

*(dd-MMM-yyyy date format, weekday-only pulls, GMP/GLP friendly)*

Stability studies are a core part of pharmaceutical development and commercial manufacturing. However, in many companies, the stability sampling schedule is still managed manually in Excel:

* Someone calculates the time points (0, 3, 6, 9, 12, 18, 24, 36 months…)
* Dates are added by hand for each batch and each storage condition
* Sampling windows (e.g., ±3 days) are calculated manually
* Weekends are often forgotten until the last minute

This manual approach is time-consuming and error-prone, especially when you have many products, batches, and storage conditions.

This article introduces a simple but powerful **Excel VBA macro** that turns Excel into a small “stability scheduler”:

* **Fully in English** (code and output)
* Date format: **`dd-MMM-yyyy`** (e.g., `05-Dec-2025`), which is unambiguous and widely used in GMP/GLP documentation
* Generates an **ICH-style** schedule for:

  * Long-term stability
  * Accelerated stability
  * Intermediate stability
* Automatically calculates:

  * Target date
  * Earliest date (Target – 3 days)
  * Latest date (Target + 3 days)
  * **Recommended pull date (Mon–Fri only)** – avoiding Saturday and Sunday
* Includes a second macro that **highlights due pulls** (rows where today is within ±3 days of a time point)

> ⚠️ This tool is a **working aid**, not a replacement for a formally approved stability protocol. It should be used alongside your SOPs and quality system.

---

## 1. Concept and Design

The macro follows a typical ICH-style design for stability testing:

* **Long-term**:
  0, 3, 6, 9, 12, 18, 24, 36 months
* **Accelerated**:
  0, 3, 6 months
* **Intermediate**:
  0, 6, 9, 12 months

For each time point, the macro calculates:

* **Target date**: Start date (T0) + N months
* **Earliest date**: Target date – 3 days
* **Latest date**: Target date + 3 days

Then, to make it more practical for real operations:

* It selects a **Recommended pull date (Mon–Fri)** within that [Earliest, Latest] window:

  * If the target date is already a weekday → use it
  * If the target date falls on Saturday/Sunday → pick the closest weekday in the window

    * First try earlier days (backwards)
    * If none, try later days (forwards)

The tool also includes a **“Check Due Samples”** macro that:

* Scans the schedule
* Highlights all rows where **today** is between Earliest and Latest date
* Shows a message listing all due pulls with their target and recommended dates

---

## 2. Preparing the Excel Workbook

### 2.1. Create the workbook and Start date cell

1. Open **Excel** and create a new workbook.
2. On **Sheet1** (or any sheet you prefer), set up the start date cell:

   * In cell **A2**, type: `Start date`
   * In cell **B2**, enter the **start date (T0)** for the batch, for example:
     `05-Dec-2025`

You can format cell B2 as **Date → dd-MMM-yyyy** if you like, but the macro will work as long as Excel recognizes it as a date.

Later, for each new batch, you will simply change B2 and re-generate the schedule.

### 2.2. Enable the Developer tab

If you don’t see the **Developer** tab in the Ribbon:

1. Go to **File → Options**.
2. Click **Customize Ribbon**.
3. On the right side, tick **Developer**.
4. Click **OK**.

You should now see a **Developer** tab at the top of Excel.

### 2.3. Save the workbook as macro-enabled

1. Go to **File → Save As**.
2. Choose your folder (e.g., a GMP-shared folder).
3. Under **Save as type**, select:
   **Excel Macro-Enabled Workbook (*.xlsm)**
4. Give it a meaningful name, for example:
   `Stability_Schedule_ICH.xlsm`
5. Click **Save**.

---

## 3. Adding the VBA Macro

### 3.1. Open the VBA Editor and create a module

1. Go to **Developer → Visual Basic** (or press `Alt + F11`).
2. In the VBA editor, choose **Insert → Module**.

   * A new module (e.g., `Module1`) will appear on the left.

### 3.2. Paste the full VBA code

Copy the entire code below and paste it into that module.

```vba
Option Explicit

Const START_DATE_CELL As String = "B2"
Const WINDOW_DAYS As Long = 3                 ' ± days around the target date
Const DATE_FMT As String = "dd-mmm-yyyy"      ' e.g. 05-Dec-2025

' Storage conditions (English text shown in sheet)
Const LONG_TERM_COND As String = "Long-term (e.g. 30°C / 75% RH)"
Const INTERMEDIATE_COND As String = "Intermediate (e.g. 30°C / 65% RH)"
Const ACCELERATED_COND As String = "Accelerated (e.g. 40°C / 75% RH)"

' Study type labels (English text shown in sheet)
Const LONG_TERM_STUDY As String = "Long-term stability"
Const INTERMEDIATE_STUDY As String = "Intermediate stability"
Const ACCELERATED_STUDY As String = "Accelerated stability"

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
               " (for example 05-Dec-2025).", vbExclamation
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
    
    ' Header row (English)
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
    nextRow = AddScheduleBlock(ws, LONG_TERM_STUDY, LONG_TERM_COND, longTermMonths, startDate, nextRow)
    
    ' Blank row between blocks
    nextRow = nextRow + 1
    
    ' Accelerated block
    nextRow = AddScheduleBlock(ws, ACCELERATED_STUDY, ACCELERATED_COND, acceleratedMonths, startDate, nextRow)
    
    ' Blank row between blocks
    nextRow = nextRow + 1
    
    ' Intermediate block
    nextRow = AddScheduleBlock(ws, INTERMEDIATE_STUDY, INTERMEDIATE_COND, intermediateMonths, startDate, nextRow)
    
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

' === Helper: label such as "12 months (1 year)" ===
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

' === Workday helpers (Mon–Fri only) ===
Private Function IsWorkday(ByVal d As Date) As Boolean
    ' Weekday with Monday as 1, Sunday as 7
    Dim wd As Long
    wd = Weekday(d, vbMonday)
    IsWorkday = (wd >= 1 And wd <= 5)   ' Monday–Friday
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
    
    ' Search backwards from target to earliest
    d = targetDate - 1
    Do While d >= earliestDate
        If IsWorkday(d) Then
            RecommendedPullDate = d
            Exit Function
        End If
        d = d - 1
    Loop
    
    ' If nothing backwards, search forwards from target to latest
    d = targetDate + 1
    Do While d <= latestDate
        If IsWorkday(d) Then
            RecommendedPullDate = d
            Exit Function
        End If
        d = d + 1
    Loop
    
    ' Fallback (rare): return target date
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
```

After pasting, you can close the VBA editor and return to Excel.

---

## 4. Running the Macro and Reading the Schedule

### 4.1. Generating the stability schedule

1. Make sure **B2** on your sheet contains a valid start date, e.g. `05-Dec-2025`.
2. Go to **Developer → Macros**.
3. Select **`CreateICHStabilitySchedule`**.
4. Click **Run**.

The macro will:

* Clear any old data in the range A1:K500.
* Create a table with headers in row 1:

| Column | Meaning                         |
| ------ | ------------------------------- |
| A      | Study type                      |
| B      | Storage condition               |
| C      | Nominal time point              |
| D      | Offset (months)                 |
| E      | Target date                     |
| F      | Earliest date (–3 days)         |
| G      | Latest date (+3 days)           |
| H      | Recommended pull date (Mon–Fri) |
| I      | Actual pull date                |
| J      | Analyst initials                |
| K      | Remarks                         |

* Add the **Long-term**, **Accelerated**, and **Intermediate** blocks one after another, with a blank row between them.
* Format columns E–I with the **dd-MMM-yyyy** pattern (e.g., `05-Dec-2025`).

### 4.2. Understanding the “Recommended pull date (Mon–Fri)”

* The **Target date** is purely theoretical (T0 + N months). It may fall on any day, including weekends.
* The **Earliest** and **Latest** dates define the acceptable sampling window (e.g., ±3 days around Target).
* The **Recommended pull date (Mon–Fri)** is the tool’s practical suggestion:

  * If the target date is a weekday and within the window → use the target date.
  * If the target date falls on Saturday or Sunday:

    * The macro first searches backwards for the closest **weekday** within the window.
    * If none exists (rare), it searches forwards.

In practice, your team would usually schedule sampling according to column **H**, and then record the real sampling date in column **I**.

---

## 5. Checking Due Samples (Daily/Weekly)

To help operations and QC keep control of upcoming pulls, use the second macro: **CheckDueSamples_ICH**.

1. Go to **Developer → Macros**.
2. Select **`CheckDueSamples_ICH`**.
3. Click **Run**.

The macro will:

* Remove any previous row highlighting in A2:K…
* Loop over all rows, and for each:

  * Check whether **TODAY** is between **Earliest date (F)** and **Latest date (G)**.
  * If yes:

    * Fill the entire row with a yellow background.
    * Add that row’s information (study type, time point, target date, recommended date) to the message text.
* At the end:

  * If at least one row is due, it shows a message like:

    > TODAY (15-Dec-2025) is within ±3 days for the following pulls:
    > Long-term stability | 12 months (1 year) | target: 14-Dec-2025 | recommended: 15-Dec-2025

  * If no row is due:

    > Today is not within ±3 days of any planned pull.

You can use this macro as a **daily or weekly check** to see which samples should be pulled.

---

## 6. Optional: Add Buttons on the Sheet

To make the tool friendly for non-technical users, you can add simple buttons on the sheet.

### 6.1. Button for “Create ICH Schedule”

1. Go to **Developer → Insert → Button (Form Control)**.
2. Click somewhere on the sheet (e.g., near B2).
3. In the **Assign Macro** dialog, choose `CreateICHStabilitySchedule` → **OK**.
4. Right-click the button → **Edit Text** → rename it to:
   **Create ICH Schedule**

### 6.2. Button for “Check Due Samples”

1. Developer → Insert → Button (Form Control).
2. Click on the sheet to place the button.
3. Assign macro: `CheckDueSamples_ICH`.
4. Edit text to: **Check Due Samples (±3 days)**.

Now users can simply click the buttons instead of navigating through the Macros dialog.

---

## 7. Customisation for Your SOP

The code is written so you can easily adapt it to your company’s procedures.

### 7.1. Change the ± window

To change from ±3 days to ±5 or ±7 days:

```vba
Const WINDOW_DAYS As Long = 3
```

Change `3` to `5`, `7`, etc.

### 7.2. Adjust time points

For example, if your long-term plan includes 48 and 60 months:

```vba
longTermMonths = Array(0, 3, 6, 9, 12, 18, 24, 36, 48, 60)
```

You can similarly adjust accelerated and intermediate arrays if your protocol defines different time points.

### 7.3. Change storage condition text

For other climatic zones or specific conditions:

```vba
Const LONG_TERM_COND As String = "Long-term (e.g. 25°C / 60% RH)"
Const ACCELERATED_COND As String = "Accelerated (e.g. 40°C / 75% RH)"
```

Replace with your actual storage conditions.

### 7.4. Change the date format (if needed)

Currently, the macro uses:

```vba
Const DATE_FMT As String = "dd-mmm-yyyy"
```

This format (e.g., `05-Dec-2025`) is highly recommended for documentation because it avoids confusion between day and month. If your local SOP requires something else (e.g., `yyyy-mm-dd`), you can modify this constant — but make sure the format is still clear and consistent.

---

## 8. GMP/GLP and Data Integrity Considerations

While this tool can dramatically reduce manual calculation errors and save time, remember:

* The macro is **not** a replacement for your **approved stability protocol**.
* Use it as a **calculation and planning aid**, and ensure that:

  * The logic has been reviewed by QA or a qualified person.
  * The file is stored in a controlled location with reasonable access control.
  * Any changes to the macro are documented or version-controlled.
  * Printed or exported schedules are clearly traceable to the protocol and batch.

If your company has a formal **computerized system validation (CSV)** or **GAMP** framework, you may choose to treat this Excel tool as a low-complexity utility and validate it accordingly (e.g., by documenting test cases and screenshots of expected schedules).

---

## 9. Conclusion

With just one Excel workbook and a few dozen lines of VBA, you can:

* Automatically generate ICH-style stability sampling schedules
* Use a clear, GMP-friendly date format (`dd-MMM-yyyy`)
* Avoid weekend pulls by using weekday-only recommended dates
* Highlight due pulls and support daily/weekly stability planning

This approach is especially useful for:

* Smaller companies that do not yet have an enterprise LIMS or stability management system
* R&D teams running many exploratory stability studies
* Sites that want a simple, transparent, and easily auditable tool to complement their formal protocols

You can freely adapt and improve the code to match your SOPs, add more storage conditions, or integrate with other Excel-based QC tools in your quality system.
