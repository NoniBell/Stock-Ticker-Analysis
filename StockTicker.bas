Attribute VB_Name = "Module1"
Sub StockTicker():


'declare and loop through worksheets
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate

'return ticker symbol
'return yearly change from beginning opening to end closing
'return percent change from beginning opening to end closing
'return total stock volume

'find the last row that has values
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'declare ticker variable
Dim ticker As String

'create variable to keep track of location in summary table
Dim sum_row As Long
sum_row = 2

'declare first opening, last closing, and percent change as variables
Dim openV As Double
Dim closeV As Double
Dim percentC As Double


'declare volume and incrementer variables
Dim volume As Long
Dim i As Long

'begin for loop to compile all variables
For i = 2 To LastRow
        'set openV and volume
        openV = Cells(i, 3).Value
        volume = 0
        'check if ticker value changes
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'set and insert ticker symbol
                ticker = Cells(i, 1).Value
                Range("I" & sum_row).Value = ticker
                
                'set closeV
                closeV = Cells(i, 6).Value
                
                'insert yearly change and percent change values
                Range("J" & sum_row).Value = closeV - openV
                
                'check if closeV or openV are equal to 0
                If closeV = 0 Or openV = 0 Then
                    percentC = 0
                Else
                    percentC = (closeV - openV) / openV * 100
                End If
            
                'insert percent and format as percent
                Range("K" & sum_row).Value = percentC
                Range("K" & sum_row).NumberFormat = "0.00%"
                
                'check and fill percent ranges if greater/less than 0
                If Range("K" & sum_row).Value < 0 Then
                    Range("K" & sum_row).Interior.ColorIndex = 3
                ElseIf Range("K" & sum_row).Value > 0 Then
                    Range("K" & sum_row).Interior.ColorIndex = 4
                End If
                
                'set and insert volume symbol
                volume = volume + Cells(i, 7).Value
                Range("L" & sum_row).Value = volume
                
                'increase sum_row
                sum_row = sum_row + 1
                
                'reset volume, ticker, percentC
                volume = 0
                ticker = ""
                percentC = 0
            Else
                'increment volume
                volume = volume + Cells(i, 7).Value
            End If
        Next i
        
        
'search completed table for greatest increase, decrease, and total volume
'create and set greatest increase, decrease, and tv variables
Dim greatestI As Double
Dim greatestD As Double
Dim greatestTV As Long

'set greatestI and greatestD as smallest qualifying values
greatestI = 0.01
greatestD = -0.01
greatestTV = 0

'declare increemnter variable
Dim j As Long

'declare most, least, and volume ticker variables
Dim increaseT As String
Dim decreaseT As String
Dim volumeT As String

'begin for loop
For j = 2 To LastRow
    'check if given cell's percent change is greater than given values
    If Range("K" & j).Value > greatestI Then
        greatestI = Range("K" & j).Value
        increaseT = Range("I" & j).Value
    ElseIf Range("K" & j).Value < greatestD Then
        greatestD = Range("K" & j).Value
        decreaseT = Range("I" & j).Value
    End If
    
    'check volumes
    If Range("L" & j).Value > greatestTV Then
        greatestTV = Range("L" & j).Value
        volumeT = Range("I" & j).Value
    End If
    
    Next j
    
    'insert final values for:
    'greatest increase, with formatting
    Range("P" & 2).Value = increaseT
    Range("Q" & 2).Value = greatestI
    Range("Q" & 2).NumberFormat = "0.00%"
    
    'greatest decrease, with formatting
    Range("P" & 3).Value = decreaseT
    Range("Q" & 3).Value = greatestD
    Range("Q" & 3).NumberFormat = "0.00%"
    
    'greatest total volume
    Range("P" & 4).Value = volumeT
    Range("Q" & 4).Value = greatestTV
    
'next sheet
     Next ws
    MsgBox ("Fixed!")
                
            

End Sub
