Attribute VB_Name = "Process"
Option Explicit

' Algorithm
'
' OpenCV-Python Tutorials > Image Processing in OpenCV > Image Thresholding
' https://docs.opencv.org/master/d7/d4d/tutorial_py_thresholding.html

Sub adaptive_threshold_mean()
    Dim image1 As Variant
    Dim image2 As Variant
    
    Dim H As Long, W As Long
    
    H = Sheets("before").Range("A1").CurrentRegion.Rows.Count
    W = Sheets("before").Range("A1").CurrentRegion.Columns.Count
    
    ReDim image2(1 To H, 1 To W) As Long
    
    
    '
    ' (1)
    '
    image1 = Sheets("before").Range("A1").Resize(H, W)
    
    
    '
    ' (2)
    '
    Dim i As Long, j As Long
    Dim s As Long, s1 As Long, s2 As Long
    Dim t As Long, t1 As Long, t2 As Long
    Dim n As Long, sum As Long
    Dim threshold As Double
    
    Const C As Long = 2
    
    For i = 1 To H
        For j = 1 To W
            s1 = max(i - 1, 1): s2 = min(i + 1, H)
            t1 = max(j - 1, 1): t2 = min(j + 1, W)
            
            n = (s2 - s1 + 1) * (t2 - t1 + 1): sum = 0
            
            For s = s1 To s2
                For t = t1 To t2
                    sum = sum + image1(s, t)
                Next t
            Next s
            
            threshold = sum / n
            
            If image1(i, j) > threshold - C Then
                image2(i, j) = 255
            Else
                image2(i, j) = 0
            End If
        Next j
    Next i
    
    
    '
    ' (3)
    '
    Sheets("after").Range("A1").Resize(H, W) = image2
End Sub

Function max(x As Long, y As Long) As Long
    max = IIf(x < y, y, x)
End Function

Function min(x As Long, y As Long) As Long
    min = IIf(x < y, x, y)
End Function

