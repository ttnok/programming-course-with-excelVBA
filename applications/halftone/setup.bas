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
        .ColumnWidth = 1 ' �Ƃ肠���� .Width  > 0  �ƂȂ�悤�ɂ��Ă����B
        .RowHeight = 1   ' �Ƃ肠���� .Height > 0  �ƂȂ�悤�ɂ��Ă����B
        
        For i = 1 To 5  ' ���񃋁[�v
            a = .ColumnWidth / .Width
            b = .RowHeight / .Height
            .Resize(H, W).ColumnWidth = pt * a
            .Resize(H, W).RowHeight = pt * b
        Next i
    End With
End Sub


