Sub start_program()
    
Dim warning_message As String
w1 = " Excel screen updating will be disabled during processing."
w2 = "The objective is to make the processing go faster, but will STILL TAKE A LONG TIME - are you sure you want to continue?"

If MsgBox(w1 + w2, vbYesNo) = vbYes Then
    'MsgBox "user clicked yes"
    'start the add_vol_for_ticker subroutine
    Call add_vol_for_ticker
    Else
    'MsgBox "user clicked no"
        'do not run the program
        Exit Sub
    End If
    



End Sub

Sub add_vol_for_ticker()

Dim ticker As String
Dim rowindex As Integer
Dim colindex  As Integer
Dim ticker_volumn As Double
ticker_volumn = 0
rowindex = 2

'start with 2016
  Worksheets("2016").Activate


  Dim ws As Worksheet
  Set ws = Worksheets("2016")
  'delete existing summary sheet from columns(i to l )
  Columns("I:L").Clear
  
  lrow = Cells(Rows.Count, 1).End(xlUp).Row
  lcol = Cells(1, Columns.Count).End(xlToLeft).Column
  
  
  'make header for results
 
        Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        Range("I1:L1").Columns.AutoFit
        Range("I1:L1").Font.Bold = True
        Range("I1:L1").Interior.ColorIndex = 6
  
  
' clear results for testing
ws.Range("I2:L21").Value = ""
 
 'Turn off screen updating to make excel go faster
 Application.ScreenUpdating = False

  For i = 2 To lrow
    'check if same
   
    If ws.Range("A" & i) = ws.Range("A" & i + 1) Or ws.Range("A" & i) <> "" And ws.Range("A" & i + 1) = "" Then
        ticker = ws.Range("A" & i).Value
        ticker_volumn = ticker_volumn + ws.Range("G" & i).Value
        
    Else
        'put the result in the sheet
        Range("I" & rowindex).Value = ticker
        Range("L" & rowindex).Value = ticker_volumn
        'add 1 to rowindex to get ready for next ticker
        rowindex = rowindex + 1
        ticker_volumn = 0
    End If
    'reached end of data put last values in result
        Range("I" & rowindex).Value = ticker
        Range("L" & rowindex).Value = ticker_volumn
    
    Next i
  
    ' now figure out yearly change, percent change for each ticker
        ' find number of rows and columns in summary section
        lrow = Range("I1").End(xlDown).Row
        lcol = ws.Range("I1:L1").Columns.Count
    
    'loop ticker to search and get beginning price and ending price for each ticker
    Dim Rng As Range
    For j = 2 To lrow
    
        ticker_to_search = ws.Range("I" & j)
        
        If Trim(ticker_to_search) <> "" Then
            With Sheets("2016").Range("A:A")
                Set Rng = .Find(What:=ticker_to_search, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
                If Not Rng Is Nothing Then
                'found first occurance of ticker_to_search
                  Application.Goto Rng, True
'                  Debug.Print (Rng.Address)
'                  Debug.Print Rng.Row
'                  Debug.Print Rng.Column
                  year_opening_price = Cells(Rng.Row, 3)
'                  Debug.Print year_opening_price
                  
'            Else
            'probably have to end loop - not sure yet
'                MsgBox "Nothing found"
            End If
        End With
        
        
    Dim Rng2 As Range
    
    If Trim(ticker_to_search) <> "" Then
        With Sheets("2016").Range("A:A")
            Set Rng2 = .Find(What:=ticker_to_search, _
                            After:=.Cells(1), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False)
            If Not Rng2 Is Nothing Then
                Application.Goto Rng2, True
'                  Debug.Print (Rng2.Address)
'                  Debug.Print Rng2.Row
'                  Debug.Print Rng2.Column
                  year_closing_price = Cells(Rng2.Row, 6)
'                  Debug.Print year_closing_price
'            Else
'                MsgBox "Nothing found"
            End If
        End With
    End If
    ws.Range("J" & j) = year_closing_price - year_opening_price
'    ws.Range("K" & j) = Round(((year_closing_price - year_opening_price) / year_opening_price) * 100, 2)
    ws.Range("K" & j) = ((year_closing_price - year_opening_price) / year_opening_price)

    
    
        
    End If
    
    Next j
    'turn screenupdating back on
    Application.ScreenUpdating = True
    'got to the top of the worksheet
    ActiveWindow.ScrollRow = 1
    
    'Apply conditional formatting red for negative numbers, green for positive, bold text
    Dim fmt_yearly_change As Range
    Set fmt_yearly_change = Range("J2", Range("J2").End(xlDown))
    fmt_yearly_change.FormatConditions.Delete
    
    'define format conditions
    Set less_than_0 = fmt_yearly_change.FormatConditions.Add(xlCellValue, xlLess, 0)
    Set greater_than_0 = fmt_yearly_change.FormatConditions.Add(xlCellValue, xlGreater, 0)
    'set fill colors and fonts
        With less_than_0
            .Interior.Color = vbRed
            .Font.Bold = True
        
        End With
        
        With greater_than_0
            .Interior.Color = vbGreen
            .Font.Bold = True
        
        End With
        
End Sub