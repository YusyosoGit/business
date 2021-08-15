VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'このコードを記載しているシートのセルに次の値を設定しておく
'B2:比較元のブック名
'B3:比較先のブック名
'比較対象の両方のブックの該当シートをあらかじめ開いておくこと。

'プログラム実行後
'値に違いがあったら、そのセルを選択して処理をストップ
'値が等しいならメッセージを表示して終了する

Public Sub CompSheets()
    'Dim path
    Dim nmSrcBook As String
    Dim nmDstBook As String
    Dim nmSrcSheet As String
    Dim nmDstSheet As String

    Dim rgDst As Range
    Dim rgSrc As Range
    Dim r As Range
    Dim cell As Range
    Dim rgDstEnd As Range
    Dim rgSrcEnd As Range
    Dim shDst As Worksheet
    Dim shSrc As Worksheet
    Dim match As Boolean
    
    'path = Range("B1").Text
    'If Right(path, 1) <> "\" Then path = path & "\"
    
    nmSrcBook = Range("B2").Text
    nmDstBook = Range("B3").Text
    'Debug.Print Workbooks(Range("B3").Text).Worksheets.Count
    
    '処理の対象はこのコードを実行しているワークブックと同じパスのエクセルファイル
    Set shDst = Workbooks(nmDstBook).ActiveSheet
    Set shSrc = Workbooks(nmSrcBook).ActiveSheet
    '範囲指定の対象のすべてに、シート名を含めないといけない
    Set rgSrcEnd = shSrc.Range(shSrc.Cells(1, 1), shSrc.Range("A1").SpecialCells(xlLastCell))
    Set rgDstEnd = shDst.Range(shDst.Cells(1, 1), shDst.Range("A1").SpecialCells(xlLastCell))
    
    Debug.Print rgSrcEnd.Rows.Count
    Debug.Print rgDstEnd.Rows.Count
    Debug.Print rgSrcEnd.Columns.Count
    Debug.Print rgDstEnd.Columns.Count
    
    
    If rgDstEnd.Rows.Count <> rgSrcEnd.Rows.Count Or _
        rgDstEnd.Columns.Count <> rgSrcEnd.Columns.Count Then
        ' https://daitaideit.com/vba-range/
        MsgBox "サイズ違い"
        Exit Sub
    End If
    
    For Each cell In rgDstEnd
        Debug.Print cell.Row & "," & cell.Column
        If cell.Value <> rgSrcEnd(cell.Row, cell.Column).Value Then
            shDst.Activate
            cell.Select
            'MsgBox "値違い"
            Exit Sub
        End If
    Next
    MsgBox _
            shDst.Name & "と、" & shSrc.Name & vbCrLf & _
            "の内容が一致しました。"
End Sub
