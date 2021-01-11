Attribute VB_Name = "Module1"
Sub ticker()

'declare variables
'Dim sht As Worksheets
'Dim r As Integer 'row iterator
'Dim c As Integer 'column iterator ---- check if needed

'variables for summary range(s)
Dim tckrsum As Range
Dim yrlychg As Range
Dim perchg As Range
Dim totvol As Range
Dim tckrinc As Range
Dim tckrdec As Range
Dim toptotvol As Range

'variables for data processing
Dim ticker As Range
Dim nxticker As Range
Dim openprc As Double
Dim clseprc As Double
Dim chgprc As Double
Dim volit As Long
Dim volume As Range
Dim thisdt As Range
Dim nxtdt As Range
Dim minyear As String
Dim maxyear As String

'start looking at sheets
For Each sht In Worksheets
'    Set sht =
    
    'find last row and last column of sheet
    lr = sht.Cells(Rows.Count, "b").End(xlUp).Row
    lc = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
    
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
        Set volume = sht.Range("G" & r)
        Set thisdt = sht.Range("B" & r)
        Set nxtdt = sht.Range("B" & nr)
                
        'get minyear
        If r = 2 Then
'            Debug.Print thisdt.Value
'            Debug.Print CStr(thisdt.Value)
'            Debug.Print "#" & Left(Right(CStr(thisdt.Value), 4), 2) & "/" & Right(CStr(thisdt.Value), 2) & "/" & Left(CStr(thisdt.Value), 4) & "#"
            
            minyear = Left(CStr(thisdt.Value), 4)
            Debug.Print minyear
            openprc = sht.Range("C" & r).Value
        End If
        
        'check year of date
        If Left(CStr(thisdt.Value), 4) <> minyear Then
            GoTo YearError
            
        
        'check ticker
        If ticker.Value <> nxticker.Value Then
           clseprc = sht.Range("f" & r).Value
        Else
            'add volume
            volit = volit + volume.Value
        End If
            
FindCloseDate:
    
    Next r
    
PrintSummary:

Next

YearError:

MsgBox ("There is more than one year in the sheet. Sub Finished")


End Sub
