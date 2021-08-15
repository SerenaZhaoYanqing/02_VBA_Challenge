Attribute VB_Name = "Module3"
Sub loopworksheets()
Dim ws As Worksheet
Application.ScreenUpdating = False
For Each ws In Worksheets
ws.Select
Call main
Next
Application.ScreenUpdating = True
End Sub

Sub main()

CreateUniqueList
lastrow = Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To lastrow
    sumvolume (i)
    yearlychange (i)
    greastvalue (i)
Next i

'Title
 Range("J1") = "Yearly change"
 Range("K1") = "Percent change"
 Range("L1") = "Total volume"
 Range("O2") = "Greatest % increase"
 Range("O3") = "Greatest % decrease"
 Range("O4") = "Greatest total stock volume"
 Range("P1") = "Ticker"
 Range("Q1") = "Value"

End Sub
Sub CreateUniqueList()

Dim lastrow As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("A1:A" & lastrow).AdvancedFilter _
    Action:=xlFilterCopy, _
    CopyToRange:=Range("I1"), _
    Unique:=True
        
End Sub

Sub sumvolume(i As Integer)

Range("L" & i).Value = WorksheetFunction.SumIf(Range("A:A"), Range("I" & i), Range("G:G"))

End Sub


Sub yearlychange(i As Integer)

'declaring variables
Dim maxdate As Long
Dim closeprice As Double
Dim mindate As Double
Dim openprice As Double

'initializing variables
maxdate = 0
mindate = 1000000000
closeprice = 0
openprice = 0.000001


lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'comparing and finding the max and min dates for specific ticker
For j = 2 To lastrow

'find max date
    If Range("A" & j).Value = Range("I" & i).Value And Range("B" & j).Value > maxdate Then
    
        maxdate = Range("B" & j).Value
        
        closeprice = Range("F" & j).Value
    End If
    
'find min date
    If Range("A" & j).Value = Range("I" & i).Value And Range("B" & j).Value < mindate Then
    
        mindate = Range("B" & j).Value
        
        openprice = Range("C" & j).Value
    End If

Next j

'çalculation
Range("J" & i).Value = closeprice - openprice
If Range("J" & i).Value > 0 Then
    Range("J" & i).Interior.ColorIndex = 4
Else
    Range("J" & i).Interior.ColorIndex = 3
End If

    
Range("K" & i).Value = (closeprice - openprice) / openprice
Range("K:K").NumberFormat = "0.00%"
End Sub

Sub greastvalue(i As Integer)
'declare variable
Dim greastincrease As Double
Dim greastdecrease As Double
Dim greaststockvolume As Double

'ínitialzing variable
greastincrease = 0
greastdecrease = 0
greaststockvolume = 0

'find greatest % increase
greastincrease = Application.WorksheetFunction.Max(Range("K:K"))
Range("Q2").Value = greastincrease
If Range("k" & i).Value = greastincrease Then
Range("P2") = Range("I" & i)
End If

'find greatest % decrease
greastdecrease = Application.WorksheetFunction.Min(Range("K:K"))
Range("Q3").Value = greastdecrease
If Range("k" & i).Value = greastdecrease Then
Range("P3") = Range("I" & i)
End If

'format percentage
Range("Q2:Q3").NumberFormat = "0.00%"

'find greatest total stock volume
greaststockvolume = Application.WorksheetFunction.Max(Range("L:L"))
Range("Q4").Value = greaststockvolume
If Range("L" & i).Value = greaststockvolume Then
Range("P4") = Range("I" & i)
End If

End Sub



