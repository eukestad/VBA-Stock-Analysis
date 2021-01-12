Attribute VB_Name = "tickeranalysis"
Sub ticker()

Application.ScreenUpdating = False
Application.Calculation = False

'declare variables
Dim sht As Worksheet


'variables for summary range(s)
Dim tckrsum As Range
Dim yrlychg As Range
Dim perchg As Range
Dim totvol As Range

'variables for bonus range(s)
Dim tckrinc As Range
Dim tckrdec As Range
Dim toptotvol As Range

'variables for data processing
Dim ticker As Range
Dim nxticker As Range
Dim prvticker As Range
Dim volume As Range
Dim openprc As Double
Dim clseprc As Double
Dim chgprc As Double
Dim chgper As Double
Dim volit As Double


'start looking at sheets
For s = 1 To Worksheets.Count
    Set sht = Sheets(s)
    sht.Activate
    
    'clear summary range contents
    slr = sht.Cells(Rows.Count, 9).End(xlUp).Row
    sht.Range("I1:Q" & slr).Delete
    
    'find last row and last column of sheet
    lr = sht.Cells(Rows.Count, 1).End(xlUp).Row
    lc = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
      
    sc = lc + 2
    sr = 1
    
    'set summary headers
    sht.Range(Cells(sr, sc), Cells(sr, sc)) = "Ticker"
    sht.Range(Cells(sr, sc + 1), Cells(sr, sc + 1)) = "Yearly Change"
    sht.Range(Cells(sr, sc + 2), Cells(sr, sc + 2)) = "Percent Change"
    sht.Range(Cells(sr, sc + 3), Cells(sr, sc + 3)) = "Total Stock Volume"

        
    'make sure the sheet is sorted
    With sht.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sht.Range("A:A"), Order:=xlAscending
        .SortFields.Add Key:=sht.Range("B:B"), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
      
    'start looking at rows
    For r = 2 To lr
        'set next row and previous row for comparisons
        nr = r + 1
        pr = r - 1 ' ---- check if needed
            
        'set ranges for values that need to be checked
        Set ticker = sht.Range("A" & r)
        Set nxticker = sht.Range("A" & nr)
        Set prvticker = sht.Range("A" & pr)
        Set volume = sht.Range("G" & r)
        
        'check ticker for start/end and get open/close price and add volume
        If ticker.Value <> prvticker.Value Then
            openprc = sht.Range("F" & r).Value
            'add volume
            volit = volit + volume.Value
        ElseIf ticker.Value <> nxticker.Value Then
            clseprc = sht.Range("f" & r).Value
            'add volume
            volit = volit + volume.Value
            'calculate yearly change
            chgprc = clseprc - openprc
            'calculate percent change
            If openprc = 0 Then
                chgper = 0
            Else
                chgper = chgprc / openprc
            End If
            
            'prepare ranges for summary
            sr = sr + 1
           
            Set tckrsum = sht.Range(Cells(sr, sc), Cells(sr, sc))
            Set yrlychg = sht.Range(Cells(sr, sc + 1), Cells(sr, sc + 1))
            Set perchg = sht.Range(Cells(sr, sc + 2), Cells(sr, sc + 2))
            Set totvol = sht.Range(Cells(sr, sc + 3), Cells(sr, sc + 3))
                       
            'print summary data and format cells
            tckrsum = ticker.Value
            yrlychg = chgprc
            If yrlychg.Value < 0 Then
                yrlychg.Interior.ColorIndex = 3
            Else
                yrlychg.Interior.ColorIndex = 4
            End If
            perchg = chgper
            perchg.NumberFormat = "0.00%"
            totvol = volit
            totvol.NumberFormat = "0"
            
            '**********-----------------------bonus ranges -------------------------*********
            
            '**********-----------------------bonus ranges -------------------------*********
            
            'autofit summary range columns to contents
            slr = sht.Cells(Rows.Count, 9).End(xlUp).Row
            sht.Range("I1:Q" & slr).Columns.AutoFit
                        
        Else
            'add volume
            volit = volit + volume.Value
        End If
       

    Next r
    
Next s

MsgBox ("Completed")

Application.ScreenUpdating = True
Application.Calculation = True

End Sub

