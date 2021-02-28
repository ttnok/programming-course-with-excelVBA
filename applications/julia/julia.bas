Attribute VB_Name = "Setup"
Option Explicit


Sub Setup()
    With Range("A1")
        square .CurrentRegion
        format .CurrentRegion
        hide_numbers .CurrentRegion
    End With
End Sub


'
'指定されたセル範囲を正方形にする
'
Sub square(canvas As Range)
    ' canvas --- セル範囲
    Dim pt As Double
    Dim i As Long
    Dim a As Double, b As Double
    
    pt = Application.CentimetersToPoints(0.4)
    
    With canvas
        .Cells(1, 1).ColumnWidth = 1 ' とりあえず .Width  > 0  となるようにしておく。
        .Cells(1, 1).RowHeight = 1    ' とりあえず .Height > 0  となるようにしておく。
        
        For i = 1 To 5  ' 数回ループ
            a = .Cells(1, 1).ColumnWidth / .Cells(1, 1).Width
            b = .Cells(1, 1).RowHeight / .Cells(1, 1).Height
            .ColumnWidth = pt * a
            .RowHeight = pt * b
        Next i
    End With
End Sub


'
' 指定された範囲のセルの条件付き書式設定
'
Sub format(canvas As Range)
    With canvas
        .FormatConditions.AddColorScale ColorScaleType:=2
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueNumber
        .FormatConditions(1).ColorScaleCriteria(1).Value = 0
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(50, 60, 80)
        
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueNumber
        .FormatConditions(1).ColorScaleCriteria(2).Value = 255
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 255)
    End With
End Sub


'
' 数値を非表示にする
'
Sub hide_numbers(canvas As Range)
    canvas.NumberFormatLocal = " "
End Sub


