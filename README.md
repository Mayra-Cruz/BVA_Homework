# BVA_Homework

Sub VBA_WallStreet_Homework()


Dim ws As Worksheet
    
    
        For Each ws In Worksheets
        ws.Activate
    
        

Dim x As Double
Dim totvol As Double
Dim i As Long
Dim LastRow As Long
       
        
        x = 2
        Cells(x, 9).Value = Cells(x, 1).Value
        
    
           LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
           If Cells(i, 1).Value = Cells(x, 9).Value Then
           totvol = totvol + Cells(i, 7).Value
        
        Else
        
           Cells(x, 10).Value = totvol
           totvol = Cells(i, 7).Value
           x = x + 1
           Cells(x, 9).Value = Cells(i, 1).Value
        
        End If
        
        Next
        
           Cells(x, 10).Value = totvol
        
    
        Next
End Sub


