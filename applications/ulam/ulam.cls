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

Const L As Long = 11 ^ 2

Sub spiral()
    Dim canvas()
    Dim n As Long
    Dim ci As Long, cj As Long
    Dim di As Long, dj As Long, di_new As Long
    Dim i As Long, j As Long
    Dim s As Long, t As Long
    Dim n_steps As Long
    
    
    ReDim canvas(-L \ 2 To L \ 2, -L \ 2 To L \ 2)
    
    ci = L \ 2 + 1: cj = L \ 2 + 1
    
    
    i = 0: j = 0
    di = 1: dj = 0
    
    n = 1: n_steps = 1
    
    canvas(i, j) = " "
    
    Do
        For s = 1 To 2
            di_new = -dj
            dj = di
            di = di_new
            
            For t = 1 To n_steps
                i = i + di
                j = j + dj
                n = n + 1
                If is_prime(n) Then
                    canvas(i, j) = n
                Else
                    canvas(i, j) = " "
                End If
                
                If n >= L ^ 2 Then Exit Do
            Next t
        Next s
        
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

Sub main()
    Debug.Print is_prime(13)
End Sub

Sub setup_canvas()
    Dim pt As Double
    Dim i As Long
    Dim a As Double, b As Double

    pt = Application.CentimetersToPoints(1)

    For i = 1 To 5
        a = Range("A1").ColumnWidth / Range("A1").Width
        b = Range("A1").RowHeight / Range("A1").Height
        Range("A1").Resize(L, L).ColumnWidth = pt * a
        Range("A1").Resize(L, L).RowHeight = pt * b
    Next i
End Sub
