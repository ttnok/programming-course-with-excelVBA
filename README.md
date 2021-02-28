# Programming course with excel VBA

プログラミング入門の題材集です。
大学の授業で利用済、あるいは利用予定です。

## 方針
* 対象者は、やっと If 文、For 文が分かりかけてきたという段階の者。
* 1 次元配列、2 次元配列の理解を目標とする。
* それに伴って、多重ループの読み書きの習熟も目標となる。
* 基本的にループは For Next 文のみにしたい。Do Loop 類の文は最低限の使用に留める。
* 変数の型は、Long、Double が基本。それ以外はなるべく使用しない。止むを得ず、Variant、Boolean、String を使うことがある。
* 基本的にサブプロシージャ 1 つで完結可能な複雑さとする。関数プロシージャを使う場合はなるべく簡単なものに留める。
* コードは数十行程度。長くても 100 行程度。
* 原則として、入力部（ワークシートから配列への入力）、処理部、出力部（配列の内容をワークシートへ書き出す）の 3 部構成とする。処理部では、ワークシートにアクセスしない。
* このリポジトリに置くアプリは作り込みすぎないようにする。なるべくシンプルにして、アイディアを提供できるようにする。


## 題材一覧

このリポジトリにアップロードしてあるものにチェックを入れています。

### 利用済
- [ ] DTMF
- [ ] string matching
- [ ] 富の分布
- [ ] 画像のフィルタ処理（畳み込み）
- [x] 画像のハーフトーン処理（誤差拡散法）
- [ ] K-means 法によるクラスタリング
- [ ] K-means 法による画像の 2 値化
- [ ] tessellation
- [ ] 一次元ランダムウォーク 

### 未利用
- [ ] DLA（Diffustion-limited aggregation、拡散律速凝集）
- [ ] CpG island
- [ ] レーベンシュタイン距離
- [ ] LCS（Longest Common Subsequence、最長共通部分列）
- [x] RGB 画像（ストライプ配列）
- [x] Mandelbrot set（マンデルブロ集合）
- [x] Julia set（ジュリア集合）
- [x] Ulam's spiral（ウラムの螺旋）
- [ ] 有理数の循環分数表示
- [ ] 迷路解法
- [ ] 勾配降下法による最適化
- [ ] Bresenham's algorithm（ブレゼンハムのアルゴリズム）
- [ ] カーネル密度推定
- [ ] 逆正弦則に従う確率分布

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
* [applications フォルダ](./applications/)の 00template2.xlsm にコード例をまとめました。
    
## VBA 言語仕様、参考書など
    
* 公式仕様書 MS-VBAL
  * https://docs.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74
     
