Attribute VB_Name = "setup"
Option Explicit

Sub square()
    Dim pt As Double
    Dim i As Long
    Dim a As Double, b As Double
    Dim W As Long, H As Long
    
    'pt= Application.CentimetersToPoints(1)
    pt = 20
        
    With Range("A1")
        H = .CurrentRegion.Rows.Count
        W = .CurrentRegion.Columns.Count
        .ColumnWidth = 1 ' とりあえず .Width  > 0  となるようにしておく。
        .RowHeight = 1   ' とりあえず .Height > 0  となるようにしておく。
        
        For i = 1 To 5  ' 数回ループ
            a = .ColumnWidth / .Width
            b = .RowHeight / .Height
            .Resize(H, W).ColumnWidth = pt * a
            .Resize(H, W).RowHeight = pt * b
        Next i
    End With
End Sub


