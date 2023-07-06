VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockanalysis()

    Dim ticker As String
    Dim openDate As String
    Dim closeDate As String
    Dim openValue As Double
    Dim closeValue As Double
    Dim vol As Long
    Dim rowCount As Long
    Dim outputRow As Long
    
    
    ticket = ""
    openDate = ""
    closeDate = ""
    openValue = 0
    closeValue = 0
    vol = 0
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    outputRow = 1
    
    Range("I:L").ClearContents
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    For Row = 2 To rowCount
        If Range("A" & Row).Value <> ticker Then
        
            ticker = Range("A" & Row).Value
            openDate = Range("B" & Row).Value
            closeDate = Range("B" & Row).Value
            openValue = Range("C" & Row).Value
            closeValue = Range("F" & Row).Value
            vol = Range("G" & Row).Value
            
        
             outputRow = outputRow + 1
             
                
        Else
            If Range("B" & Row).Value < openDate Then
                openDate = Range("B" & Row).Value
                openValue = Range("C" & Row).Value
            End If
            
            If Range("B" & Row).Value > closeDate Then
                closeDate = Range("B" & Row).Value
                closeValue = Range("F" & Row).Value
            End If
            
            vol = vol + Range("G" & Row).Value
        End If
        
        Range("I" & outputRow).Value = ticker
        Range("J" & outputRow).Value = Round(closeValue - openValue, 2)
        Range("K" & outputRow).Value = Round((closeValue - openValue) / openValue, 2)
        Range("L" & outputRow).Value = vol
        Range("K" & outputRow).NumberFormat = "0.00%"
    Next Row
End Sub

