# CSV出力機能

ワークシートの可視セルを CSV ファイルへ出力する。

## 構成

| ファイル | 責務 |
|---|---|
| `Svc_CsvExport.bas` | 可視セルの抽出、エスケープ、ファイル出力 |

## 仕様

- 出力対象は `UsedRange` の可視セルのみ。非表示行(オートフィルタの除外行を含む)と非表示列は出力しない
- 文字コードは Shift-JIS 固定
- 改行コードは CRLF
- 値にカンマ・引用符・改行が含まれる場合は RFC4180 準拠でエスケープする
- セルの表示形式は反映されない(値をそのまま出力)。日付のみ `yyyy/mm/dd` に整形する
- 指数表記(`1.23E+15`)は展開して出力する

## 呼び出し側の実装例

Form 側にファイル選択とエラー表示を置く。Service は MsgBox を出さない。

```vb
Private Sub btnCsv_Click()
    On Error GoTo ErrHandler

    Dim path As String
    path = 保存先を選択(shtList.Name & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".csv")
    If path = "" Then Exit Sub          ' キャンセル

    Application.ScreenUpdating = False
    Dim cnt As Long
    cnt = Svc_CsvExport.可視セル出力(shtList, path)
    Application.ScreenUpdating = True

    If cnt = 0 Then
        MsgBox "出力対象のデータがありません。", vbExclamation
    Else
        MsgBox cnt & "件を出力しました。" & vbCrLf & path, vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "CSV出力に失敗しました。" & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "エラー"
End Sub

Private Function 保存先を選択(ByVal defaultName As String) As String
    Dim v As Variant
    v = Application.GetSaveAsFilename( _
            InitialFileName:=defaultName, _
            FileFilter:="CSVファイル (*.csv), *.csv", _
            Title:="CSVの保存先を指定してください")
    If VarType(v) = vbBoolean Then Exit Function   ' キャンセル
    保存先を選択 = CStr(v)
End Function
```

## 設計上の判断

### 標準機能(SaveAs xlCSV)を使わなかった理由

- **可視セルのみの出力ができない。** シート全体が出力されるため、フィルタで絞った状態が反映されない
- **文字コード・書式を制御できない。** 先頭 0 が落ちる、数値が指数表記になる、日付形式が環境依存になる
- **副作用が大きい。** 新規ブックが開き、アクティブブックが切り替わる

### 速度面の考慮

- セル単位のアクセスを避け、`Range.Value` で範囲を一括取得する
- 文字列連結は `Join` を使う(逐次連結は行数に対して二乗で遅くなる)
- ファイル書き込みは 1 回にまとめる

数万行でも実用的な速度で動作する。

## 既知の制約

- `SpecialCells(xlCellTypeVisible)` は範囲を Area に分割する。フィルタで飛び飛びになるほど Area 数が増え、ループ回数が増加する
- セルの表示形式(通貨記号、パーセント)は反映されない。表示通りに出力する場合は `.Text` を参照する必要があるが、セル単位アクセスとなり大幅に遅くなる
- `UsedRange` は書式のみ設定されたセルも含むため、意図より広い範囲が対象になる場合がある
