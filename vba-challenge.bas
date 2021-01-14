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
Dim tckrtpvol As Range
Dim incchgper As Range
Dim decchgper As Range
Dim gtotvol As Range

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

Dim maxticker As String
Dim maxperchg As Double
maxperchg = 0
Dim minticker As String
Dim minperchg As Double
minperchg = 0
Dim volticker As String
Dim volvalue As Double
volvalue = 0

'start looking at sheets
For s = 1 To Worksheets.Count
    Set sht = Sheets(s)
    sht.Activate
    
    'clear summary range contents
    slr = sht.Cells(Rows.Count, 9).End(xlUp).Row
    sht.Range("I1:Q" & slr).Delete
    
    'make sure the sheet is sorted
    With sht.Sort
        .SortFields.Clear
        .SortFields.Add Key:=sht.Range("A:A"), Order:=xlAscending
        .SortFields.Add Key:=sht.Range("B:B"), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    'find last row and last column of sheet
    lr = sht.Cells(Rows.Count, 1).End(xlUp).Row
    lc = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
      
    'set summary headers
    Dim colhead As Range
    sr = 1
    
    For sc = (lc + 2) To (lc + 6)
'        Debug.Print sc
        Set colhead = sht.Cells(sr, sc)
'        Debug.Print sc - (lc + 2) + 1
        If sc - (lc + 2) + 1 = 1 Then
            colhead = "Ticker"
        ElseIf sc - (lc + 2) + 1 = 2 Then
            colhead = "Yearly Change ($)"
        ElseIf sc - (lc + 2) + 1 = 3 Then
            colhead = "Percent Change"
        ElseIf sc - (lc + 2) + 1 = 4 Then
            colhead = "Total Stock Volume"
        End If
        
    Next sc
    
    sc = (lc + 2)
    
'    sht.Range("I1") = "Ticker"
'    sht.Range("J1") = "Yearly Change"
'    sht.Range("K1") = "Percent Change"
'    sht.Range("L1") = "Total Stock Volume"
    
    'set bonus ranges
    sht.Range("N2") = "Greatest % Increase"
    sht.Range("N3") = "Greatest % Decrease"
    sht.Range("N4") = "Greatest Total Volume"
    sht.Range("O1") = "Ticker"
    sht.Range("P1") = "Value"
    Set tckrinc = sht.Range("O2")
    Set tckrdec = sht.Range("O3")
    Set tckrtpvol = sht.Range("O4")
    Set incchgper = sht.Range("P2")
    Set decchgper = sht.Range("P3")
    Set gtotvol = sht.Range("P4")
    
       
    'start looking at rows
    For r = 2 To lr
        'set next row and previous row for comparisons
        nr = r + 1
        pr = r - 1
            
        'set ranges for values that need to be checked
        Set ticker = sht.Range("A" & r)
        Set nxticker = sht.Range("A" & nr)
        Set prvticker = sht.Range("A" & pr)
        Set volume = sht.Range("G" & r)
        
        'check ticker for start/end and get open/close price and add volume
        If ticker.Value <> prvticker.Value Then
            
            'check for 0 value in open price and exit sub if 0 is found. otherwise set open price
            If sht.Range("F" & r).Value <> 0 Then
                openprc = sht.Range("F" & r).Value
            Else
                MsgBox ("Open Price is 0, please validate cell at Column F, Row " & r & " and re-run.")
'                sht.Range("F" & r).Activate
'                GoTo exitsub
            End If
            
            'add volume
            volit = volit + volume.Value
            
        ElseIf ticker.Value <> nxticker.Value Then
            'set close price
            clseprc = sht.Range("f" & r).Value
            'add volume
            volit = volit + volume.Value
            'calculate yearly change
            chgprc = clseprc - openprc
            'calculate percent change
            If openprc <> 0 Then
                chgper = chgprc / openprc
            Else
                chgper = 0
            End If
            'bonus calculations
            If chgper > maxperchg Then
                maxperchg = chgper
                maxticker = ticker.Value
            End If
            If chgper < minperchg Then
                minperchg = chgper
                minticker = ticker.Value
            End If
            If volit > volvalue Then
                volvalue = volit
                volticker = ticker.Value
            End If
            
            'prepare ranges for summary
            sr = sr + 1
            
            For sc = (lc + 2) To (lc + 6)
'                Debug.Print sc
'                Debug.Print sc - (lc + 2) + 1
                If sc - (lc + 2) + 1 = 1 Then
                    Set tckrsum = sht.Cells(sr, sc)
                ElseIf sc - (lc + 2) + 1 = 2 Then
                    Set yrlychg = sht.Cells(sr, sc)
                ElseIf sc - (lc + 2) + 1 = 3 Then
                    Set perchg = sht.Cells(sr, sc)
                ElseIf sc - (lc + 2) + 1 = 4 Then
                    Set totvol = sht.Cells(sr, sc)
                End If
                
            Next sc
            
            sc = (lc + 2)
                                   
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
                                   
        Else
            'add volume
            volit = volit + volume.Value
        End If
       

    Next r
    
        
    'print max summary data (bonus)
    tckrinc = maxticker
    tckrdec = minticker
    tckrtpvol = volticker
    incchgper = maxperchg
    incchgper.NumberFormat = "0.00%"
    decchgper = minperchg
    decchgper.NumberFormat = "0.00%"
    gtotvol = volvalue
    gtotvol.NumberFormat = "0"
    
    'autofit summary range columns to contents
    slr = sht.Cells(Rows.Count, 9).End(xlUp).Row
    sht.Range("I1:Q" & slr).Columns.AutoFit
    
Next s


'exitsub:


MsgBox ("Completed")

Application.ScreenUpdating = True
Application.Calculation = True

End Sub

