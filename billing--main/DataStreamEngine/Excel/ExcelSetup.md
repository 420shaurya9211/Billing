# Excel Dashboard Auto-Refresh Setup

## Quick Setup (3 minutes)

### Step 1: Open Excel and Create New Workbook
1. Open Excel → **File → New → Blank Workbook**
2. Save as **`Dashboard.xlsm`** (Macro-Enabled Workbook) in the `output` folder

### Step 2: Import CSV Data via Power Query

For each CSV file (`customers.csv`, `invoices.csv`, `loans.csv`):

1. **Data → Get Data → From File → From Text/CSV**
2. Browse to `output/customers.csv`
3. Click **Load** → Loads into a named table
4. Repeat for `invoices.csv` and `loans.csv`

### Step 3: Add VBA Auto-Refresh Macro

1. Press **Alt + F11** to open the VBA editor
2. Double-click **ThisWorkbook** in the Project Explorer
3. Paste this code:

```vba
' ═══════════════════════════════════════════════════
'  Auto-Refresh Timer for Real-Time CSV Streaming
'  Configurable interval (default: 5 seconds)
' ═══════════════════════════════════════════════════

Private Const REFRESH_INTERVAL_SECONDS As Long = 5
Private NextRefreshTime As Date
Private IsRunning As Boolean

Private Sub Workbook_Open()
    StartAutoRefresh
End Sub

Public Sub StartAutoRefresh()
    IsRunning = True
    ScheduleNextRefresh
    Application.StatusBar = "Auto-Refresh: ON (every " & REFRESH_INTERVAL_SECONDS & "s)"
End Sub

Public Sub StopAutoRefresh()
    IsRunning = False
    On Error Resume Next
    Application.OnTime NextRefreshTime, "ThisWorkbook.RefreshAllData", , False
    On Error GoTo 0
    Application.StatusBar = "Auto-Refresh: OFF"
End Sub

Private Sub ScheduleNextRefresh()
    If Not IsRunning Then Exit Sub
    NextRefreshTime = Now + TimeSerial(0, 0, REFRESH_INTERVAL_SECONDS)
    Application.OnTime NextRefreshTime, "ThisWorkbook.RefreshAllData"
End Sub

Public Sub RefreshAllData()
    If Not IsRunning Then Exit Sub
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    ' Refresh all Power Query connections
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.OLEDBConnection.BackgroundQuery = False
        conn.Refresh
    Next conn
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Last refresh: " & Format(Now, "hh:mm:ss") & " | Next: " & Format(NextRefreshTime + TimeSerial(0, 0, REFRESH_INTERVAL_SECONDS), "hh:mm:ss")
    
    ScheduleNextRefresh
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = "Refresh error at " & Format(Now, "hh:mm:ss") & ": " & Err.Description
    ScheduleNextRefresh ' Keep trying
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    StopAutoRefresh
    Application.StatusBar = False
End Sub
```

4. Press **Ctrl+S** to save, close VBA editor

### Step 4: Create Dashboard Charts

1. Go to the **customers** sheet → Select data → **Insert → Pivot Table**
2. Create pivot charts for:
   - Customer count by city
   - Total purchases distribution
   - Active loans breakdown
3. Repeat for invoices (revenue by metal type, GST summary) and loans (status breakdown)

### Step 5: Test It

1. Run the C# DataStreamEngine (`dotnet run`)
2. Open `Dashboard.xlsm` — it will auto-start refreshing
3. Watch the data update live every 5 seconds!

### Controls
- To **stop** auto-refresh: Open VBA editor → Run `StopAutoRefresh`
- To **restart**: Run `StartAutoRefresh`
- Change `REFRESH_INTERVAL_SECONDS` to adjust the polling rate
