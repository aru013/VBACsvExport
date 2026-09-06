Attribute VB_Name = "Svc_CsvExport"
'==============================================================================
' Svc_CsvExport.bas
' 責務  : ワークシートの可視セルを CSV ファイルへ出力する
' 依存  : なし(シート / ファイル操作のみ)
' 禁止  : MsgBox の表示、DB アクセス
' 文字コード : Shift-JIS 固定
'
' 呼び出し例:
'   Dim cnt As Long
'   cnt = Svc_CsvExport.可視セル出力(shtList, "C:\temp\out.csv")
'==============================================================================
Option Explicit

Private Const DELIM     As String = ","
Private Const NEW_LINE  As String = vbCrLf

'------------------------------------------------------------------------------
' 可視セルのみを CSV へ出力する
'   非表示行(フィルタ除外を含む)および非表示列は出力対象外。
'   セルの表示形式は反映されず、値がそのまま出力される(日付のみ書式指定)。
'
'   ws       : 対象シート
'   filePath : 出力先フルパス
'   戻り値   : 出力した行数。対象データが無い場合は 0
'------------------------------------------------------------------------------
Public Function 可視セル出力(ByVal ws As Worksheet, _
                             ByVal filePath As String) As Long

    If ws Is Nothing Then
        Err.Raise vbObjectError + 700, "Svc_CsvExport.可視セル出力", _
                  "出力対象のシートが指定されていません。"
    End If
    If Trim$(filePath) = "" Then
        Err.Raise vbObjectError + 701, "Svc_CsvExport.可視セル出力", _
                  "出力先が指定されていません。"
    End If

    Dim rngAll As Range
    Set rngAll = ws.UsedRange
    If rngAll Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(rngAll) = 0 Then Exit Function

    '--- 可視セル範囲(行方向のフィルタ結果を反映)---
    Dim rngVisible As Range
    On Error Resume Next
    Set rngVisible = rngAll.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rngVisible Is Nothing Then Exit Function

    '--- 可視「列」は SpecialCells で除外されないため別途判定する ---
    Dim visCols() As Long
    Dim colCount As Long
    colCount = 可視列取得(rngAll, visCols)
    If colCount = 0 Then Exit Function

    '--- 出力バッファ ---
    Dim buf() As String
    ReDim buf(0 To rngAll.Rows.Count - 1)
    Dim bufIdx As Long

    Dim line() As String
    ReDim line(0 To colCount - 1)

    '--- Area 単位で配列化して処理する ---
    Dim area As Range
    For Each area In rngVisible.Areas

        Dim data As Variant
        data = 範囲を配列化(area)

        Dim firstCol As Long
        firstCol = area.Column

        Dim r As Long, c As Long, relCol As Long
        For r = 1 To UBound(data, 1)
            For c = 0 To colCount - 1
                relCol = visCols(c) - firstCol + 1
                If relCol >= 1 And relCol <= UBound(data, 2) Then
                    line(c) = CSVエスケープ(data(r, relCol))
                Else
                    line(c) = ""
                End If
            Next c
            buf(bufIdx) = Join(line, DELIM)
            bufIdx = bufIdx + 1
        Next r
    Next area

    If bufIdx = 0 Then Exit Function
    ReDim Preserve buf(0 To bufIdx - 1)

    '--- 一括書き込み(行ごとの I/O を避ける)---
    Call ファイル書込(filePath, Join(buf, NEW_LINE) & NEW_LINE)

    可視セル出力 = bufIdx
End Function

'------------------------------------------------------------------------------
' 可視列の列番号を配列で返す
'   outCols : 可視列の列番号(0 起算の配列に列番号を格納)
'   戻り値  : 可視列数
'------------------------------------------------------------------------------
Private Function 可視列取得(ByVal rngAll As Range, _
                            ByRef outCols() As Long) As Long
    Dim tmp() As Long
    ReDim tmp(0 To rngAll.Columns.Count - 1)

    Dim n As Long, i As Long
    For i = 1 To rngAll.Columns.Count
        If Not rngAll.Columns(i).EntireColumn.Hidden Then
            tmp(n) = rngAll.Columns(i).Column
            n = n + 1
        End If
    Next i

    If n > 0 Then
        ReDim Preserve tmp(0 To n - 1)
        outCols = tmp
    End If
    可視列取得 = n
End Function

'------------------------------------------------------------------------------
' 範囲を 2 次元配列で返す
'   Range.Value は単一セルの場合スカラーを返すため、常に 2 次元へ揃える
'------------------------------------------------------------------------------
Private Function 範囲を配列化(ByVal rng As Range) As Variant
    If rng.Cells.Count = 1 Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = rng.Value
        範囲を配列化 = one
    Else
        範囲を配列化 = rng.Value
    End If
End Function

'------------------------------------------------------------------------------
' CSV の値をエスケープする
'   - 引用符は 2 つ重ねる
'   - 区切り文字 / 引用符 / 改行を含む場合は全体を引用符で囲む
'   - Null / Empty / エラー値は空文字とする
'------------------------------------------------------------------------------
Private Function CSVエスケープ(ByVal v As Variant) As String
    Dim s As String

    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function

    Select Case VarType(v)
        Case vbDate
            s = Format$(v, "yyyy/mm/dd")
        Case vbDouble, vbSingle, vbCurrency, vbDecimal
            ' 指数表記を避けて展開する
            If v = Int(v) And Abs(v) < 1E+15 Then
                s = Format$(v, "0")
            Else
                s = Format$(v, "0.##############")
            End If
        Case Else
            s = CStr(v)
    End Select

    If InStr(s, """") > 0 Then s = Replace$(s, """", """""")

    If InStr(s, DELIM) > 0 _
       Or InStr(s, """") > 0 _
       Or InStr(s, vbLf) > 0 _
       Or InStr(s, vbCr) > 0 Then
        s = """" & s & """"
    End If

    CSVエスケープ = s
End Function

'------------------------------------------------------------------------------
' ファイルへ書き込む(Shift-JIS)
'   既存ファイルは上書きする
'------------------------------------------------------------------------------
Private Sub ファイル書込(ByVal filePath As String, ByVal content As String)
    Dim fno As Integer
    fno = FreeFile

    Open filePath For Output As #fno
    Print #fno, content;
    Close #fno
End Sub
