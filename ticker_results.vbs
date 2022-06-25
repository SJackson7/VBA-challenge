Attribute VB_Name = "Module1"
Sub ticker()
'create to allow worksheet loop
Dim tickerYear As Worksheet

'set ticker name variable
Dim tickerName As String

'set range of, first and last row varible for each ticker to grab intitial open and lase close values
Dim tStartRow As Long
Dim tLastRow As Long
tStartRow = 0
tLastRow = 0

'set year-beginning (yb) opening price variable and initialize to 0
Dim openVal As Double
openVal = 0

'set year-end (ye) closing price variable and initialize to 0
Dim closeVal As Double
closeVal = 0

'set yearly change variable and initialize to 0
Dim yr_chg As Double
yr_chg = 0

'set percent change variable and initalize to 0
Dim pct_chg As Double
pct_chg = 0

'set total stock variable and initialize to 0
Dim ttlVol As LongLong
ttlVol = 0

'start row location for each ticker in summary table
Dim result_table_row As Integer
result_table_row = 2

'make sure first sheet "2018" is active
If ActiveSheet.Name <> "2018" Then
    Worksheets("2018").Activate
End If

'loop to run combined data for each sheet (ticker year)
For Each tickerYear In ThisWorkbook.Worksheets
    sLastRow = tickerYear.Cells(Rows.Count, 1).End(xlUp).Row 'get last row on sheet
    'add ticker name column
    tickerYear.Cells(1, 9).Value = "Ticker"
    'add metric columns
    tickerYear.Cells(1, 10).Value = "Yearly Change"
    tickerYear.Cells(1, 11).Value = "Percent Change"
    tickerYear.Cells(1, 12).Value = "Total Stock Volume"
    'loop through all ticker symbols
    For i = 2 To sLastRow
        If tickerYear.Cells(i + 1, 1).Value <> tickerYear.Cells(i, 1).Value Then   'validate if checking same ticker symbol from previous
            tickerName = tickerYear.Cells(i, 1).Value                              'set ticker symbol name
            
            'get first row number of ticker symbol for yb opening value
            tStartRow = WorksheetFunction.Match(tickerName, tickerYear.Range("A1:A" & sLastRow), 0)

            openVal = tickerYear.Cells(tStartRow, 3).Value              'get ticker yb opening value, rounded to 2 digits due to floating point issue in DOUBLE data type
            closeVal = tickerYear.Cells(i, 6).Value                     'get ticker ye closing value based on last instance of ticker symbol,, rounded to 2 digits due to floating point issue in DOUBLE data type
            ttlVol = ttlVol + tickerYear.Cells(i, 7).Value              'add to ticker total volume
            yr_chg = closeVal - openVal                      'get open/close diff
            pct_chg = (closeVal - openVal) / openVal         'get open/close diff%
            tickerYear.Range("i" & result_table_row).Value = tickerName 'print ticker symbol name to result table
            tickerYear.Range("l" & result_table_row).Value = ttlVol     'print total volume of ticker to result table
            tickerYear.Range("j" & result_table_row).Value = yr_chg     'print ticker open/close diff
            tickerYear.Range("k" & result_table_row).Value = pct_chg    'print ticker open/close diff%
            'debug test openVal show:Range("m" & result_table_row).Value = openVal
            'debug test closeVal show:Range("n" & result_table_row).Value = closeVal
            
            'set number format and color for numerical change results, highlighting row for better readability
            tickerYear.Range("j" & result_table_row).NumberFormat = "+0.00;-0.00;±0.00"
            tickerYear.Range("k" & result_table_row).NumberFormat = "+0.00%;-0.00%;±0.00%"
            If yr_chg < 0 Then
                tickerYear.Range(tickerYear.Cells(result_table_row, "i"), tickerYear.Cells(result_table_row, "L")).Interior.Color = RGB(255, 199, 206)
                tickerYear.Range(tickerYear.Cells(result_table_row, "i"), tickerYear.Cells(result_table_row, "L")).Font.Color = RGB(156, 0, 6)
            ElseIf yr_chg > 0 Then
                tickerYear.Range(tickerYear.Cells(result_table_row, "i"), tickerYear.Cells(result_table_row, "L")).Interior.Color = RGB(198, 239, 206)
                tickerYear.Range(tickerYear.Cells(result_table_row, "i"), tickerYear.Cells(result_table_row, "L")).Font.Color = RGB(0, 97, 0)
            Else
                tickerYear.Range(tickerYear.Cells(result_table_row, "i"), tickerYear.Cells(result_table_row, "L")).Interior.ColorIndex = 0
                tickerYear.Range(tickerYear.Cells(result_table_row, "i"), tickerYear.Cells(result_table_row, "L")).Font.ColorIndex = 0
            End If
            
            'increment result table row
            result_table_row = result_table_row + 1
            
            'reinitialize variables to 0 for next ticker name
            tickerName = vbNullString
            ttlVol = 0
            openVal = 0
            closeVal = 0
            tStartRow = 0
            tLastRow = 0
            yr_chg = 0
            pct_chg = 0
        
        'add to ticker symbol total if same name
        Else
            ttlVol = ttlVol + tickerYear.Cells(i, 7).Value
        End If
    Next i
    
    'functionality to return stock with greatest % increse, greatest % decrease and greatest total volume
    'same sub used

    'add new row and column headers
    tickerYear.Cells(2, 15).Value = "Greatest % Incr."
    tickerYear.Cells(3, 15).Value = "Greatest % Decr."
    tickerYear.Cells(4, 15).Value = "Greatest Total Vol."
    tickerYear.Cells(1, 16).Value = "Ticker"
    tickerYear.Cells(1, 17).Value = "Value"
    
    'get greatest% increase/decrease and greatest value via max/min functions
    rLastRow = tickerYear.Cells(Rows.Count, 9).End(xlUp).Row    'get last row from results
    Dim gpi As Double                                           'set var for greatest% incr.
    Dim gpd As Double                                           'set var for greatest% decr.
    Dim gtv As LongLong                                         'set var for greatest total vol.
   
    gpi = WorksheetFunction.Max(tickerYear.Range("K2:K" & rLastRow))    'largest incr%
    gpd = WorksheetFunction.Min(tickerYear.Range("K2:K" & rLastRow))    'largest decr%
    gtv = WorksheetFunction.Max(tickerYear.Range("L2:L" & rLastRow))    'largest total vol
    
    'temp vars to hold ticker name row number matching incr/decr/highest vol
    ticker_gpi = WorksheetFunction.Match(gpi, tickerYear.Range("K1:K" & rLastRow), 0)
    ticker_gpd = WorksheetFunction.Match(gpd, tickerYear.Range("K1:K" & rLastRow), 0)
    ticker_gtv = WorksheetFunction.Match(gtv, tickerYear.Range("L1:L" & rLastRow), 0)
        
    'set number format for %s and values and print results
    tickerYear.Cells(2, 17).NumberFormat = "+0.00%;-0.00%;±0.00%"
    tickerYear.Cells(3, 17).NumberFormat = "+0.00%;-0.00%;±0.00%"
    tickerYear.Cells(2, 16).Value = tickerYear.Range("I" & ticker_gpi).Value
    tickerYear.Cells(2, 17).Value = gpi
    tickerYear.Cells(3, 16).Value = tickerYear.Range("I" & ticker_gpd).Value
    tickerYear.Cells(3, 17).Value = gpd
    tickerYear.Cells(4, 16).Value = tickerYear.Range("I" & ticker_gtv).Value
    tickerYear.Cells(4, 17).Value = gtv
    
    'reinitialize for next sheet loop and auto-fit result columns
    tickerYear.Columns("I:L").AutoFit
    tickerYear.Columns("O:Q").AutoFit
    i = 0
    tickerName = vbNullString
    ttlVol = 0
    openVal = 0
    closeVal = 0
    tStartRow = 0
    tLastRow = 0
    yr_chg = 0
    pct_chg = 0
    result_table_row = 2
    gpi = 0
    gpd = 0
    gtv = 0
    ticker_gpi = 0
    ticker_gpd = 0
    ticker_gtv = 0
    rLastRow = 0
   
Next tickerYear

End Sub

