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

Const L As Long = 11 ^ 2   'Šï”

Sub spiral()
    Dim canvas() As Variant
    ReDim canvas(-L \ 2 To L \ 2, -L \ 2 To L \ 2)
    
    Dim n As Long
    Dim i As Long, j As Long
    Dim di As Long, dj As Long, di_new As Long
    
    Dim n_steps As Long
    Dim edge As Long, s As Long
    
    
    Range("A1").Resize(L, L).ClearContents

    
    n = 1
    i = 0: j = 0
    di = 1: dj = 0
'    canvas(i, j) = " "
    
    
    n_steps = 1
    Do
        For edge = 1 To 2
            di_new = -dj
            dj = di
            di = di_new
            
            For s = 1 To n_steps
                i = i + di: j = j + dj
                n = n + 1
                
                If is_prime(n) Then
                    canvas(i, j) = n
                Else
'                    canvas(i, j) = " "
                End If
                
                If n >= L ^ 2 Then Exit Do
            Next s
        Next edge
        
        n_steps = n_steps + 1
    Loop
    
    
    Range("A1").Resize(L, L) = canvas
End Sub

Function is_prime(n As Long) As Boolean
    Dim i As Long
    
    i = 2
    For i = 2 To n - 1
        If n Mod i = 0 Then Exit For
    Next i
    
    is_prime = i = n
End Function

Sub select_canvas()
    Range("A1").Resize(L, L).Select
End Sub
