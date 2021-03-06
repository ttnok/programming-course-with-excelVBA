Attribute VB_Name = "process"
Sub halftone()
    Dim H As Long, W As Long
    Dim image As Variant
    Dim i As Long, j As Long
    Dim brightness As Long, r As Long
    
    H = Sheets("original").Range("A1").CurrentRegion.Rows.Count
    W = Sheets("original").Range("A1").CurrentRegion.Columns.Count
    
    '
    ' (1)
    '
    image = Sheets("original").Range("A1").Resize(H, W)
    
    '
    ' (2)
    '
    For i = 1 To H
        For j = 1 To W
            brightness = image(i, j)
            
            If brightness <= 127 Then
                r = brightness
                brightness = 0
            Else
                r = brightness - 255
                brightness = 255
            End If
            
            image(i, j) = brightness
            
            If j <> W Then
                image(i, j + 1) = image(i, j + 1) + (5 / 16) * r
            End If
            
            If i <> H Then
                image(i + 1, j) = image(i + 1, j) + (5 / 16) * r
                
                If j <> 1 Then
                    image(i + 1, j - 1) = image(i + 1, j - 1) + (3 / 16) * r
                End If
                
                
                If j <> W Then
                    image(i + 1, j + 1) = image(i + 1, j + 1) + (3 / 16) * r
                End If
            End If
        Next j
    Next i
    
    '
    ' (3)
    '
    Sheets("halftoned").Range("A1").Resize(H, W) = image
End Sub
