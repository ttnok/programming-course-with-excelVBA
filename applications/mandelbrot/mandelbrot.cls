VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub mandelbrot()
    Const I_MAX As Long = 500
    Const J_MAX As Long = 500
    Const X_MIN As Double = -1#, X_MAX As Double = 2.2
    Const Y_MIN As Double = -1.6, Y_MAX As Double = 1.6
    Const ITER_MAX As Long = 100
    
    Dim screen() As Variant
    Dim i As Long, j As Long
    Dim iter As Long
    Dim c1 As Double, c2 As Double
    Dim z1 As Double, z2 As Double, z1_temp As Double
    Dim x As Double, y As Double
    Dim hx As Double, hy As Double
    
    
    Range("A1").Resize(I_MAX, J_MAX).ClearContents
    
    screen = Range("A1").Resize(I_MAX, J_MAX)
    
    For i = 1 To I_MAX
        For j = 1 To J_MAX
            screen(i, j) = 0
        Next j
    Next i
    
    hx = (X_MAX - X_MIN) / (J_MAX - 1)
    hy = (Y_MAX - Y_MIN) / (I_MAX - 1)
    
    For i = 1 To I_MAX
        c2 = Y_MAX - (i - 1) * hy
        
        For j = 1 To J_MAX
            c1 = X_MIN + (j - 1) * hx
            
            z1 = 0: z2 = 0
            
            For iter = 1 To ITER_MAX
                z1_temp = z1 ^ 2 - z2 ^ 2 - c1
                z2 = 2 * z1 * z2 - c2
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
