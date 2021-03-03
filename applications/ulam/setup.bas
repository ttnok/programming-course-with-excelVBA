Attribute VB_Name = "Setup"
Option Explicit

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

