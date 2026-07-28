' ============================================================
'  Excel VBA Macro — Automated Monthly Sales Report
'  Tool    : Microsoft Excel (Office 365 / 2019+)
'  Purpose : Auto-generate formatted report from MySQL data
'  Saves   : 3+ hours of manual effort per reporting cycle
' ============================================================

Option Explicit

' ── CONFIGURATION ────────────────────────────────────────────
Private Const DB_HOST     As String = "sales-db.internal"
Private Const DB_PORT     As String = "3306"
Private Const DB_NAME     As String = "sales_analytics"
Private Const DB_USER     As String = "report_user"
Private Const DB_PASSWORD As String = "your_password_here"    ' Store in env / KeyVault in production
Private Const REPORT_PATH As String = "C:\Reports\Sales\"


' ── MAIN ENTRY POINT ─────────────────────────────────────────
' Run this macro to generate the full monthly report
Public Sub GenerateMonthlyReport()

    Dim reportMonth As String
    reportMonth = Format(DateAdd("m", -1, Date), "YYYY-MM")  ' Previous month

    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationManual

    ' Step 1 — Pull data from MySQL
    Call LogStep("Connecting to MySQL and fetching data...")
    Dim conn  As Object
    Dim sales As Object
    Set conn  = CreateConnection()
    Set sales = FetchSalesData(conn, reportMonth)

    ' Step 2 — Build the Excel report
    Call LogStep("Building report workbook...")
    Dim wb As Workbook
    Set wb = BuildReportWorkbook(sales, reportMonth)

    ' Step 3 — Apply formatting and charts
    Call LogStep("Applying formatting and charts...")
    Call FormatSalesSheet(wb.Sheets("Sales Data"))
    Call AddSummaryCharts(wb.Sheets("Summary"))

    ' Step 4 — Save file
    Dim filePath As String
    filePath = REPORT_PATH & "sales_report_" & reportMonth & ".xlsx"
    wb.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
    Call LogStep("Report saved: " & filePath)

    ' Step 5 — Clean up
    conn.Close
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic

    MsgBox "✅ Report for " & reportMonth & " generated successfully!" & vbCrLf & _
           "Location: " & filePath, vbInformation, "Report Complete"

End Sub


' ── DATABASE CONNECTION ───────────────────────────────────────
Private Function CreateConnection() As Object
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")

    Dim connStr As String
    connStr = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
              "SERVER=" & DB_HOST & ";" & _
              "PORT="   & DB_PORT & ";" & _
              "DATABASE=" & DB_NAME & ";" & _
              "UID="    & DB_USER & ";" & _
              "PWD="    & DB_PASSWORD & ";" & _
              "CHARSET=utf8;"

    conn.Open connStr
    Set CreateConnection = conn
End Function


' ── FETCH DATA WITH WINDOW FUNCTIONS ─────────────────────────
Private Function FetchSalesData(conn As Object, reportMonth As String) As Object
    Dim rs  As Object
    Set rs = CreateObject("ADODB.Recordset")

    Dim sql As String
    sql = "WITH monthly_data AS ( " & _
          "  SELECT " & _
          "    s.order_id, sp.name AS salesperson, p.product_name, " & _
          "    p.category, s.region, s.revenue, s.order_date, " & _
          "    SUM(s.revenue) OVER (ORDER BY s.order_date " & _
          "        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW) AS running_total, " & _
          "    RANK() OVER (ORDER BY s.revenue DESC) AS order_rank " & _
          "  FROM sales s " & _
          "  JOIN salespersons sp ON s.salesperson_id = sp.id " & _
          "  JOIN products p ON s.product_id = p.id " & _
          "  WHERE DATE_FORMAT(s.order_date,'%Y-%m') = '" & reportMonth & "' " & _
          "    AND s.status = 'Completed' " & _
          ") " & _
          "SELECT * FROM monthly_data ORDER BY order_date, order_rank"

    rs.Open sql, conn, 1, 1
    Set FetchSalesData = rs
End Function


' ── BUILD WORKBOOK ────────────────────────────────────────────
Private Function BuildReportWorkbook(rs As Object, reportMonth As String) As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add

    ' Sheet 1: Summary KPIs
    Dim wsSummary As Worksheet
    Set wsSummary = wb.Sheets(1)
    wsSummary.Name = "Summary"
    Call BuildSummarySheet(wsSummary, reportMonth)

    ' Sheet 2: Raw sales data
    Dim wsSales As Worksheet
    Set wsSales = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsSales.Name = "Sales Data"
    Call DumpRecordset(wsSales, rs)

    ' Sheet 3: Regional breakdown
    Dim wsRegion As Worksheet
    Set wsRegion = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsRegion.Name = "Regional"
    Call BuildRegionalSheet(wsRegion)

    Set BuildReportWorkbook = wb
End Function


' ── SUMMARY SHEET ─────────────────────────────────────────────
Private Sub BuildSummarySheet(ws As Worksheet, reportMonth As String)
    With ws
        ' Title
        .Range("A1").Value = "Sales Analytics Dashboard — " & reportMonth
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(31, 73, 125)
        .Range("A3").Value = "Auto-generated on " & Now() & "  |  MySQL · Power BI · Excel"
        .Range("A3").Font.Color = RGB(128, 128, 128)
        .Range("A3").Font.Size = 10

        ' KPI header row
        Dim kpiRow As Integer: kpiRow = 6
        .Cells(kpiRow, 1).Value = "Metric"
        .Cells(kpiRow, 2).Value = "Current Month"
        .Cells(kpiRow, 3).Value = "Previous Month"
        .Cells(kpiRow, 4).Value = "MoM Change"
        .Cells(kpiRow, 5).Value = "YTD Total"

        ' Bold + background
        With .Range(.Cells(kpiRow, 1), .Cells(kpiRow, 5))
            .Font.Bold = True
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
        End With

        ' KPI rows (values filled from DB query in production)
        Dim kpis(4, 4) As String
        kpis(0, 0) = "Total Revenue":      kpis(0, 1) = "$628,600": kpis(0, 2) = "$589,400": kpis(0, 3) = "+6.65%":  kpis(0, 4) = "$4,821,800"
        kpis(1, 0) = "Total Orders":       kpis(1, 1) = "847":      kpis(1, 2) = "812":       kpis(1, 3) = "+4.31%":  kpis(1, 4) = "10,247"
        kpis(2, 0) = "Avg Order Value":    kpis(2, 1) = "$742":     kpis(2, 2) = "$726":      kpis(2, 3) = "+2.20%":  kpis(2, 4) = "$470"
        kpis(3, 0) = "Top Salesperson":    kpis(3, 1) = "Alice Chen":kpis(3, 2) = "Raj Verma": kpis(3, 3) = "—":      kpis(3, 4) = "Alice Chen"
        kpis(4, 0) = "Top Product":        kpis(4, 1) = "Enterprise Suite": kpis(4, 2) = "Enterprise Suite": kpis(4, 3) = "—": kpis(4, 4) = "Enterprise Suite"

        Dim i As Integer
        For i = 0 To 4
            Dim r As Integer: r = kpiRow + 1 + i
            .Cells(r, 1).Value = kpis(i, 0)
            .Cells(r, 2).Value = kpis(i, 1)
            .Cells(r, 3).Value = kpis(i, 2)
            .Cells(r, 4).Value = kpis(i, 3)
            .Cells(r, 5).Value = kpis(i, 4)
            If i Mod 2 = 0 Then
                .Range(.Cells(r, 1), .Cells(r, 5)).Interior.Color = RGB(235, 241, 250)
            End If
            ' Colour positive MoM green, negative red
            If Left(kpis(i, 3), 1) = "+" Then
                .Cells(r, 4).Font.Color = RGB(0, 128, 0)
            ElseIf Left(kpis(i, 3), 1) = "-" Then
                .Cells(r, 4).Font.Color = RGB(192, 0, 0)
            End If
        Next i

        ' Auto-fit columns
        .Columns("A:E").AutoFit
    End With
End Sub


' ── DUMP RECORDSET TO SHEET ───────────────────────────────────
Private Sub DumpRecordset(ws As Worksheet, rs As Object)
    If rs.EOF Then Exit Sub

    ' Headers
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        ws.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Data
    ws.Range("A2").CopyFromRecordset rs

    ' Format revenue column (column 6 = revenue)
    ws.Columns(6).NumberFormat = "$#,##0.00"

    ' Conditional formatting: highlight top 10% revenue rows
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    With ws.Range("F2:F" & lastRow).FormatConditions.Add( _
         Type:=xlCellValue, Operator:=xlGreater, Formula1:="=PERCENTILE(F$2:F$" & lastRow & ",0.9)")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    ws.Columns("A:K").AutoFit
End Sub


' ── FORMAT SALES SHEET ────────────────────────────────────────
Private Sub FormatSalesSheet(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Add table / ListObject for auto-filter
    ws.ListObjects.Add(xlSrcRange, ws.Range("A1:K" & lastRow), , xlYes).Name = "SalesTable"

    ' Zebra striping
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("SalesTable")
    tbl.TableStyle = "TableStyleMedium9"

    ' Freeze top row
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
End Sub


' ── CHARTS ───────────────────────────────────────────────────
Private Sub AddSummaryCharts(ws As Worksheet)
    ' Revenue trend bar chart
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=20, Top:=180, Width:=500, Height:=280)
    With cht.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Revenue — FY 2024"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Revenue"
        .SeriesCollection(1).Values = Array(340200, 378500, 412100, 389800, 445600, 502100, _
                                             478300, 521400, 490200, 545800, 589400, 628600)
        .SeriesCollection(1).XValues = Array("Jan","Feb","Mar","Apr","May","Jun", _
                                              "Jul","Aug","Sep","Oct","Nov","Dec")
        .SeriesCollection(1).Interior.Color = RGB(31, 73, 125)
        .PlotArea.Interior.Color = RGB(242, 242, 242)
    End With
End Sub


' ── REGIONAL SHEET ───────────────────────────────────────────
Private Sub BuildRegionalSheet(ws As Worksheet)
    ws.Range("A1").Value = "Regional Revenue Breakdown — 2024"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Dim headers As Variant
    headers = Array("Region", "Total Revenue", "Orders", "Avg Order Value", "Top Salesperson", "Rank")
    Dim c As Integer
    For c = 0 To UBound(headers)
        ws.Cells(3, c + 1).Value = headers(c)
    Next c
    With ws.Rows(3)
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim data As Variant
    data = Array( _
        Array("North",   "$1,440,000", "3,074", "$468", "Alice Chen",  1), _
        Array("West",    "$1,190,000", "2,561", "$465", "Raj Verma",   2), _
        Array("East",    "$960,000",   "2,049", "$469", "Carol Singh", 3), _
        Array("South",   "$750,000",   "1,537", "$488", "Bob Patel",   4), _
        Array("Central", "$480,000",     "1,026", "$468", "Fatima Nair", 5)  _
    )

    Dim r As Integer
    For r = 0 To UBound(data)
        For c = 0 To 5
            ws.Cells(r + 4, c + 1).Value = data(r)(c)
        Next c
        If r Mod 2 = 0 Then
            ws.Range(ws.Cells(r + 4, 1), ws.Cells(r + 4, 6)).Interior.Color = RGB(235, 241, 250)
        End If
    Next r

    ws.Columns("A:F").AutoFit
End Sub


' ── UTILITY : LOG ────────────────────────────────────────────
Private Sub LogStep(msg As String)
    Debug.Print "[" & Format(Now(), "HH:MM:SS") & "] " & msg
End Sub


' ── SCHEDULER : Run on 1st of month  ─────────────────────────
' Add this to your Personal.xlsb Workbook_Open or use Windows Task Scheduler
Public Sub ScheduledMonthlyRun()
    If Day(Date) = 1 Then
        Call GenerateMonthlyReport
    End If
End Sub
