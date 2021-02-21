# Programming course with excel VBA

## memo

* データ入力時に `Debug.Assert` で確認
* データ入力時に、ブレークポイントとローカルウインドウで確認
* ウォッチ式の利用
* If と Stop の利用


## セルをピクセルとして利用する際の Tips

* 一斉書き込み
  * `Range("...").Resize(H, W) =` two-dim-array
* ユーザ定義書式（`" "`）で文字を非表示
* 相対アドレス指定
  * `Range("....").Cells(i, j)`

* セルを正方形にする
	```bas
  Sub main()
  	Dim pt As Double
  	Dim i As Long
  	Dim a As Double, b As Double
    
  	pt = Application.CentimetersToPoints(1)
    
  	For i = 1 To 5
    	a = Range("A1").ColumnWidth / Range("A1").Width
    	b = Range("A1").RowHeight / Range("A1").Height
    	Range("A1").Resize(100, 100).ColumnWidth = pt * a
    	Range("A1").Resize(100, 100).RowHeight = pt * b
  	Next i
	End Sub
	```
