
# Hướng dẫn tạo lịch lấy mẫu độ ổn định theo ICH bằng Excel Macro

*(Định dạng ngày dd-MMM-yyyy, tránh cuối tuần, phục vụ GMP/GLP)*

Trong thực tế sản xuất và R&D, việc lập lịch lấy mẫu độ ổn định (stability sampling schedule) thường được làm thủ công trên Excel: tự tính các mốc 0, 3, 6, 9, 12, 18, 24, 36 tháng, tự cộng ngày, tự kẻ bảng, rất dễ sai sót, nhất là khi có nhiều sản phẩm và nhiều điều kiện bảo quản.

Theo ICH Q1A(R2) và các guideline tương tự, lịch lấy mẫu độ ổn định thường gồm:

* **Long-term**: mỗi 3 tháng trong năm 1 (0, 3, 6, 9, 12),
  mỗi 6 tháng trong năm 2 (18, 24),
  và hàng năm sau đó (36, 48, 60…).
* **Accelerated**: tối thiểu 0, 3, 6 tháng.
* **Intermediate**: khi cần, thường dùng các mốc 0, 6, 9, 12 tháng.

Về **Good Documentation Practice (GDP)**, nhiều tài liệu khuyến cáo dùng định dạng ngày **dd-MMM-yyyy** (ví dụ 14-DEC-2020) để tránh nhầm lẫn giữa dạng châu Âu (dd/mm/yyyy) và dạng Mỹ (mm/dd/yyyy).

Bài viết này giới thiệu một **Excel VBA macro**:

* Code hoàn toàn **tiếng Anh** (dễ chia sẻ, audit, review).
* Kết quả trên sheet cũng **tiếng Anh**, dễ dùng cho hồ sơ song ngữ.
* Ngày hiển thị theo dạng **dd-MMM-yyyy** (GMP-friendly, tránh mơ hồ).
* Tự động tạo lịch:

  * Long-term, Accelerated, Intermediate theo mốc kiểu ICH (0–36 tháng).
  * Cửa sổ ±3 ngày quanh Target date (Earliest / Latest).
  * Thêm cột **Recommended pull date (Mon–Fri)**: tự tránh Thứ 7 & Chủ nhật.
* Có macro thứ hai để **tô màu vàng các mốc đang “đến hạn”** (hôm nay nằm trong ±3 ngày).

---

## 1. Ý tưởng tổng quan

Macro này biến Excel thành một “mini-stability scheduler”:

1. Bạn nhập **Start date (T0)** vào ô B2 (ví dụ ngày sản xuất hoặc ngày bắt đầu ổn định).
2. Bấm chạy macro **CreateICHStabilitySchedule**:

   * Excel tự tạo bảng:

     * **Long-term stability**: 0, 3, 6, 9, 12, 18, 24, 36 months.
     * **Accelerated stability**: 0, 3, 6 months.
     * **Intermediate stability**: 0, 6, 9, 12 months.
   * Mỗi mốc có:

     * Target date (theo T0 + số tháng).
     * Earliest date (Target – 3 days).
     * Latest date (Target + 3 days).
     * Recommended pull date: chọn một ngày **thứ 2–6** nằm trong khoảng trên.
3. Hàng ngày/tuần, bạn chạy macro **CheckDueSamples_ICH**:

   * Các dòng có **TODAY** nằm trong khoảng [Earliest, Latest] sẽ được **tô vàng**.
   * Một popup hiện ra liệt kê các mốc đang đến hạn.

---

## 2. Chuẩn bị file Excel

### 2.1. Tạo file & ô Start date

1. Mở **Excel**, tạo một workbook mới.
2. Trên **Sheet1**:

   * Ô **A2**: gõ chữ `Start date`
   * Ô **B2**: nhập ngày bắt đầu (T0), ví dụ:

     * `05-Dec-2025`
     * hoặc `05/12/2025` rồi format sau đó.
3. Sau này mỗi lô bạn chỉ cần đổi **B2** và chạy lại macro.

### 2.2. Bật tab Developer

Nếu chưa thấy tab **Developer**:

1. Vào **File → Options**.
2. Chọn **Customize Ribbon**.
3. Bên phải, tick chọn **Developer** → **OK**.

### 2.3. Lưu file dạng macro-enabled

1. Vào **File → Save As**.
2. Chọn nơi lưu (Desktop, thư mục dự án…).
3. Ở **Save as type**, chọn:

> **Excel Macro-Enabled Workbook (*.xlsm)**

4. Đặt tên, ví dụ: `Stability_Schedule_ICH.xlsm` → **Save**.

---

## 3. Thêm VBA macro vào workbook

### 3.1. Mở cửa sổ VBA

1. Vào tab **Developer → Visual Basic**
   (hoặc nhấn phím tắt **Alt + F11**).
2. Trong cửa sổ VBA:

   * Menu **Insert → Module**
   * Xuất hiện `Module1` (hoặc tên tương tự).

### 3.2. Dán toàn bộ code dưới đây

Copy **nguyên khối** code (từ `Option Explicit` đến hết file) vào Module vừa tạo:

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

Sau khi dán xong, đóng cửa sổ VBA (dấu X), quay lại Excel.

---

## 4. Cách chạy macro & đọc kết quả

### 4.1. Tạo lịch ổn định ICH

1. Đảm bảo **ô B2** đã có ngày Start date, ví dụ `05-Dec-2025`.
2. Vào tab **Developer → Macros**.
3. Chọn macro **`CreateICHStabilitySchedule`** → bấm **Run**.

Kết quả:

* Vùng **A1:K…** sẽ có bảng:

| Col | Nội dung                        |
| --- | ------------------------------- |
| A   | Study type                      |
| B   | Storage condition               |
| C   | Nominal time point              |
| D   | Offset (months)                 |
| E   | Target date                     |
| F   | Earliest date (-3 days)         |
| G   | Latest date (+3 days)           |
| H   | Recommended pull date (Mon–Fri) |
| I   | Actual pull date                |
| J   | Analyst initials                |
| K   | Remarks                         |

* Long-term block: 0 → 36 months.
* Một dòng trống.
* Accelerated block: 0, 3, 6 months.
* Một dòng trống.
* Intermediate block: 0, 6, 9, 12 months.

Tất cả ngày (cột E–H) được format **`dd-MMM-yyyy`**, ví dụ `05-Dec-2025`, tuân theo khuyến cáo GDP về tránh mơ hồ ngày/tháng.

### 4.2. Ý nghĩa cột Recommended pull date (Mon–Fri)

* **Target date** (E): ngày danh nghĩa, đúng theo T0 + số tháng. Có thể rơi vào thứ 7 hoặc CN.
* **Earliest / Latest** (F, G): cửa sổ ±3 ngày (theo ví dụ trong macro).
* **Recommended pull date** (H):

  * Nếu Target là ngày làm việc (Mon–Fri) và nằm trong [Earliest, Latest] → dùng luôn Target.
  * Nếu Target rơi vào T7/CN:

    * Macro sẽ ưu tiên tìm **ngày làm việc gần nhất về phía trước** trong khoảng.
    * Nếu không có, sẽ tìm **ngày làm việc gần nhất về phía sau** trong khoảng.

Khi lấy mẫu thật, bạn sẽ điền **Actual pull date** (I) – thường trùng với cột Recommended, trừ khi công ty có lý do đặc biệt.

---

## 5. Kiểm tra các mốc “đến hạn” hằng ngày

Macro thứ hai, **CheckDueSamples_ICH**, dùng để hỗ trợ QA/QC kiểm soát lịch:

1. Vào tab **Developer → Macros**.
2. Chọn **`CheckDueSamples_ICH`** → **Run**.

Macro sẽ:

* Xóa màu nền cũ trong A2:K…
* Dò từng dòng, nếu **TODAY** (ngày hiện tại của hệ thống) nằm trong khoảng [Earliest, Latest]:

  * Tô **màu vàng** cho dòng đó.
  * Thêm dòng vào message text.
* Nếu có mốc đến hạn → hiện popup, ví dụ:

> TODAY (15-Dec-2025) is within ±3 days for the following pulls:
> Long-term stability | 12 months (1 year) | target: 14-Dec-2025 | recommended: 15-Dec-2025

* Nếu không có mốc nào:

  > Today is not within ±3 days of any planned pull.

Bạn có thể chạy macro này **mỗi sáng** như một checklist nhanh cho stability.

---

## 6. Thêm nút bấm trên sheet (không bắt buộc nhưng rất tiện)

Để dễ sử dụng cho nhiều người trong nhà máy:

### 6.1. Nút “Create ICH Schedule”

1. Vào **Developer → Insert → Button (Form Control)**.
2. Click một vị trí trên sheet (ví dụ cạnh ô B2).
3. Excel hỏi **Assign Macro** → chọn `CreateICHStabilitySchedule` → OK.
4. Click phải lên nút → **Edit Text** → sửa thành:
   **Create ICH Schedule**

### 6.2. Nút “Check Due Samples”

Làm tương tự:

1. Developer → Insert → Button.
2. Assign Macro → chọn `CheckDueSamples_ICH`.
3. Sửa text thành **Check Due Samples (±3 days)**.

Từ nay người vận hành chỉ cần click nút, không cần vào menu Macros.

---

## 7. Tùy chỉnh cho SOP nội bộ

Bạn có thể chỉnh code rất dễ:

1. **Đổi khoảng ± ngày** (ví dụ ±5 hoặc ±7 ngày):

```vba
Const WINDOW_DAYS As Long = 3
```

→ đổi `3` thành `5` hoặc `7`.

2. **Đổi mốc thời gian**:

* Long-term (thêm 48, 60 tháng):

```vba
longTermMonths = Array(0, 3, 6, 9, 12, 18, 24, 36, 48, 60)
```

* Accelerated / Intermediate cũng chỉnh tương tự nếu SOP yêu cầu khác.

3. **Đổi mô tả điều kiện bảo quản** (zone khác):

```vba
Const LONG_TERM_COND As String = "Long-term (e.g. 25°C / 60% RH)"
Const ACCELERATED_COND As String = "Accelerated (e.g. 40°C / 75% RH)"
```

4. **Đổi định dạng ngày** (nếu công ty chọn dạng khác):

```vba
Const DATE_FMT As String = "dd-mmm-yyyy"
```

Có thể đổi sang `"yyyy-mm-dd"` nếu muốn dạng ISO (nhưng dd-MMM-yyyy vẫn là lựa chọn rất rõ ràng trong GDP).

---

## 8. Lưu ý GMP/GLP & Data Integrity

* Đây chỉ là **công cụ hỗ trợ tính toán** và lập lịch.

* Để dùng trong hệ thống chính thức, nên:

  * Đưa vào quản lý tài liệu (document control).
  * Xem xét **validation** của file Excel (ít nhất là kiểm tra lại logic, công thức, bảo vệ sheet, tránh chỉnh sửa nhầm code).
  * Quy định rõ trong **SOP ổn định**:

    * Định dạng ngày sử dụng.
    * Cách tính ± ngày chấp nhận được.
    * Cách lưu trữ file (backup, version, quyền chỉnh sửa).

* Macro này không thay thế cho **stability protocol** đã được phê duyệt, mà chỉ là **working tool** giúp:

  * Tính lịch nhanh và nhất quán.
  * Tránh sót mốc, tránh nhầm lẫn ngày, nhất là khi có nhiều sản phẩm & nhiều điều kiện.

---
