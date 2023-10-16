VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SummarizeTickers():
    Dim TICKER As String
    Dim PR_OPEN As Double
    Dim PR_CLOSE As Double
    'figuring that the volume needed to be a LongLong instead of a regular Long was downright painful,
    'this one issue cost me several hours of time and a considerable portion of my sanity
    'Big thanks to the Excel Error popup for providing functionally zero help whatsoever to debug this one.
    'A stacktrace or even just line number would have been really, really, really nice.
    Dim VOL As LongLong
    Dim UNIQUE_ROW As Long
    Dim ROW_CUR As Long
    Dim TICK_BIG_W As String
    Dim TICK_BIG_L As String
    Dim TICK_BIG_V As String
    Dim AMT_BIG_W As Double
    Dim AMT_BIG_L As Double
    Dim AMT_BIG_V As LongLong
    Dim WS As Worksheet
    Dim headers() As Variant
    Dim ws_counter As Long
    
    '"hardcoded" values for columns ticker, open, close, volume
    Dim COL_TIC As Long
    Dim COL_OP As Long
    Dim COL_CL As Long
    Dim COL_VOL As Long
    COL_TIC = CLng(1)
    COL_OP = CLng(3)
    COL_CL = CLng(6)
    COL_VOL = CLng(7)
    'column that calculation results *START* on
    Dim COL_RES As Long
    COL_RES = CLng(9)
    
    'This saves an astounding amount of time, and I am surprised it was not suggested for students to include this in their macros.
    Application.ScreenUpdating = False
    
    ws_counter = 0
    For Each WS In Worksheets
        WS.Range("I1:Q1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value")
        WS.Range("O2").Value = "Greatest % Increase"
        WS.Range("O3").Value = "Greatest % Decrease"
        WS.Range("O4").Value = "Greatest Total Volume"
        
        'these all need to be reset for every worksheet
        ROW_CUR = CLng(2)
        UNIQUE_ROW = CLng(2)
        TICK_BIG_W = ""
        TICK_BIG_L = ""
        TICK_BIG_V = ""
        AMT_BIG_W = 0
        AMT_BIG_L = 0
        AMT_BIG_V = CLngLng(0)
        'needs to be set this once, though it is a bit redundant
        TICKER = WS.Cells(ROW_CUR, COL_TIC).Value
        
        While TICKER <> ""
            PR_OPEN = WS.Cells(ROW_CUR, COL_OP).Value
            VOL = CLngLng(0)
            'checks if current row ticker matches next ticker,
            'advances row and adds volume to total if answer is yes
            While TICKER = WS.Cells(ROW_CUR + CLng(1), COL_TIC).Value
                VOL = VOL + CLngLng(WS.Cells(ROW_CUR, COL_VOL).Value)
                ROW_CUR = ROW_CUR + CLng(1)
            Wend
            'establish closing price for year
            PR_CLOSE = WS.Cells(ROW_CUR, COL_CL).Value
            'collect final volume for total
            VOL = VOL + CLngLng(WS.Cells(ROW_CUR, COL_VOL).Value)
            'start adding unique values to their spot at the right of the data
            WS.Cells(UNIQUE_ROW, COL_RES).Value = TICKER
            'set Yearly Change value
            WS.Cells(UNIQUE_ROW, COL_RES + 1) = PR_CLOSE - PR_OPEN
            'format Yearly Change cell and Pct change cell green if yearly change is positive, red if negative
            If WS.Cells(UNIQUE_ROW, COL_RES + 1).Value > 0 Then
                WS.Cells(UNIQUE_ROW, COL_RES + 1).Interior.ColorIndex = 4
                WS.Cells(UNIQUE_ROW, COL_RES + 2).Interior.ColorIndex = 4
            ElseIf WS.Cells(UNIQUE_ROW, COL_RES + 1).Value < 0 Then
                WS.Cells(UNIQUE_ROW, COL_RES + 1).Interior.ColorIndex = 3
                WS.Cells(UNIQUE_ROW, COL_RES + 2).Interior.ColorIndex = 3
            End If
            'set percent change. formula *should* be correct
            WS.Cells(UNIQUE_ROW, COL_RES + 2).Value = (PR_CLOSE - PR_OPEN) / PR_OPEN
            'set volume total
            WS.Cells(UNIQUE_ROW, COL_RES + 3).Value = VOL
            
            'iterate unique row forwards
            '(and iterate current row to the next one with new ticker)
            UNIQUE_ROW = UNIQUE_ROW + CLng(1)
            ROW_CUR = ROW_CUR + CLng(1)
            TICKER = WS.Cells(ROW_CUR, COL_TIC).Value
        Wend

        'for loop to calculate biggest winner, loser, and volume from summarized cells
        For i = 2 To UNIQUE_ROW
            If WS.Cells(i, COL_RES + 2).Value > AMT_BIG_W Then
                TICK_BIG_W = WS.Cells(i, COL_RES).Value
                AMT_BIG_W = WS.Cells(i, COL_RES + 2).Value
            End If
            If WS.Cells(i, COL_RES + 2).Value < AMT_BIG_L Then
                TICK_BIG_L = WS.Cells(i, COL_RES).Value
                AMT_BIG_L = WS.Cells(i, COL_RES + 2).Value
            End If
            If CLngLng(WS.Cells(i, COL_RES + 3).Value) > AMT_BIG_V Then
                TICK_BIG_V = WS.Cells(i, COL_RES).Value
                AMT_BIG_V = CLngLng(WS.Cells(i, COL_RES + 3).Value)
            End If
        Next i

        'set all cells using data gathered in for loop
        WS.Cells(2, COL_RES + 7) = TICK_BIG_W
        WS.Cells(2, COL_RES + 8) = AMT_BIG_W
        WS.Cells(3, COL_RES + 7) = TICK_BIG_L
        WS.Cells(3, COL_RES + 8) = AMT_BIG_L
        WS.Cells(4, COL_RES + 7) = TICK_BIG_V
        WS.Cells(4, COL_RES + 8) = AMT_BIG_V
        
        'reset superlatives for next WS iteration
        TICK_BIG_W = ""
        TICK_BIG_L = ""
        TICK_BIG_V = ""
        AMT_BIG_W = 0
        AMT_BIG_L = 0
        AMT_BIG_V = 0
        
        'formatting and cleanup
        ws_counter = ws_counter + 1
        WS.Columns("J").NumberFormat = "#,##0.00_);-#,##0.00 "
        WS.Columns("K").NumberFormat = "0.00%"
        WS.Range("Q2:Q3").NumberFormat = "0.00%"
        WS.Range("Q4").NumberFormat = "0.00E+0"
        WS.Columns("I:Q").AutoFit
    Next WS
    MsgBox ("Done. Processed " + Str(ws_counter) + " worksheets.")
    Application.ScreenUpdating = True
End Sub

Sub restoreVisible()
    'this was primarily used during debugging to fix the screen updates not getting re-enabled
    Application.ScreenUpdating = True
End Sub

