Attribute VB_Name = "Module1"
Option Explicit

Const W As Long = 320
Const H As Long = 240

Sub stripe()
    Dim R, G, B, RGB
    ReDim RGB(1 To H, 1 To 3 * W)
    Dim i As Long, j As Long
    
    R = Worksheets("Red").Range("A1").Resize(H, W)
    G = Worksheets("Green").Range("A1").Resize(H, W)
    B = Worksheets("Blue").Range("A1").Resize(H, W)
        
    For i = 1 To H
        For j = 1 To W
            RGB(i, 3 * j - 2) = R(i, j)
            RGB(i, 3 * j - 1) = G(i, j)
            RGB(i, 3 * j) = B(i, j)
        Next j
    Next i
    
    Sheets("RGB").Range("A1").Resize(H, 3 * W) = RGB
End Sub
