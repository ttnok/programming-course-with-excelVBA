# Programming course with excel VBA

プログラミング入門の題材集です。
大学の授業で利用済、あるいは利用予定です。

## 題材一覧

このリポジトリにアップロードしてあるものにチェックを入れています。

### 利用済
- [ ] string matching
- [ ] 富の分布
- [ ] 画像のフィルタ処理（畳み込み）
- [ ] 画像のハーフトーン処理
- [ ] K-means 法によるクラスタリング
- [ ] K-means 法による画像の 2 値化
- [ ] tessellation (affine 変換）
- [ ] 一次元ランダムウォーク 

### 未利用
- [ ] DLA
- [x] mandelbrot set
- [x] Ulam's spiral
- [ ] レーベンシュタイン距離
- [ ] 迷路解法
- [ ] 勾配降下法による最適化
- [ ] RGB 画像（ストライプ配列）
- [ ] Bresenham's algorithm
- [ ] カーネル密度推定
- [ ] 有理数の循環分数表示

## memo

* データ入力時に `Debug.Assert` で確認
* データ入力時に、ブレークポイントとローカルウインドウで確認
* ウォッチ式の利用
* If と Stop の利用


## セルをピクセルとして利用する際の Tips

* 一斉書き込み
  * `Range("...").Resize(H, W) =` two-dim-array
* ユーザ定義書式（`" "`）で数値を非表示
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
    
## VBA 言語仕様、参考書など
    
* 公式仕様書 MS-VBAL
  * https://docs.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74
     
