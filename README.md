# Programming course with excel VBA

## memo

* データ入力時に `Debug.Assert` で確認
* データ入力時に、ブレークポイントとローカルウインドウで確認


## セルをピクセルとして利用する際の Tips

* 一斉書き込み
  * `Range("...").Resize(H, W) =` two-dim-array
* ユーザ定義書式（`" "`）で文字を非表示
* 相対アドレス指定
  * `Range("....").Cells(i, j)`
