VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub julia()
    Const I_MAX As Long = 1000
    Const J_MAX As Long = 1000
    Const X_MIN As Double = -2, X_MAX As Double = 2
    Const Y_MIN As Double = -2, Y_MAX As Double = 2
    Const ITER_MAX As Long = 1000
    
    Dim screen() As Variant
    Dim i As Long, j As Long
    Dim iter As Long
    
    Dim c1 As Double, c2 As Double
    Dim z1 As Double, z2 As Double, z1_temp As Double
    Dim hx As Double, hy As Double
    
    
    Application.ScreenUpdating = False
    Range("A1").CurrentRegion.ClearContents
        
    ReDim screen(1 To I_MAX, 1 To J_MAX)
    
    
    hx = (X_MAX - X_MIN) / (J_MAX - 1)
    hy = (Y_MAX - Y_MIN) / (I_MAX - 1)
    
    c1 = -0.4
    c2 = 0.6
    
    For i = 1 To I_MAX
        
        For j = 1 To J_MAX
            z1 = X_MIN + (j - 1) * hx
            z2 = Y_MAX - (i - 1) * hy
            
            For iter = 1 To ITER_MAX
                z1_temp = z1 ^ 2 - z2 ^ 2 + c1
                z2 = 2 * z1 * z2 + c2
                z1 = z1_temp
                
                If z1 ^ 2 + z2 ^ 2 > 4 Then
                    Exit For
                End If
            Next iter
            
            screen(i, j) = iter
        Next j
    Next i
    
    
    Range("A1").Resize(I_MAX, J_MAX) = screen
End Sub


