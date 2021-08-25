Option Explicit

'このコードを記載しているシートのセルに次の値を設定しておく
'B2:比較元のブック名
'B3:比較先のブック名
'比較対象の両方のブックの該当シートをあらかじめ開いておくこと。

'プログラム実行後
'値に違いがあったら、そのセルを選択して処理をストップ
'値が等しいならメッセージを表示して終了する

Public Sub CompSheets()
    Dim path
    Dim nmSrcBook As String
    Dim nmDstBook As String
    Dim nm As String
    Dim cnt As Integer
    Dim endA

    Dim cell As Range
    Dim rgDst As Range
    Dim rgSrc As Range
    Dim shDst As Worksheet
    Dim shSrc As Worksheet
    Dim match As Boolean
    Dim ans
    
    nmSrcBook = Range("B2").Text
    nmDstBook = Range("B3").Text
    'Debug.Print Workbooks(Range("B3").Text).Worksheets.Count
    
    '処理の対象はこのコードを実行しているワークブックと同じパスのエクセルファイル
    'path = hisWorkbook.path & "\"
    nm = Dir(path & nmSrcBook)
    If nm <> "" Then
        Set shSrc = Workbooks(nm).ActiveSheet
    Else
        MsgBox "コピー元ファイルなし"
        Exit Sub
    End If
    nm = Dir(path & nmDstBook)
    If nm <> "" Then
        Set shDst = Workbooks(nm).ActiveSheet
    Else
        MsgBox "コピー先ファイルなし"
        Exit Sub
    End If
    
    Set rgSrc = shSrc.Range(shSrc.Cells(1, 1), shSrc.Range("A1").SpecialCells(xlLastCell))
    Set rgDst = shDst.Range(shDst.Cells(1, 1), shDst.Range("A1").SpecialCells(xlLastCell))
    
    Debug.Print rgSrc.Rows.Count
    Debug.Print rgDst.Rows.Count
    Debug.Print rgSrc.Columns.Count
    Debug.Print rgDst.Columns.Count
    
    
'    If rgDst.Rows.Count <> rgSrc.Rows.Count Or _
'        rgDst.Columns.Count <> rgSrc.Columns.Count Then
'        ans = MsgBox("表のサイズが異なりますが、続行しますか？", vbOKCancel)
'        If ans <> vbOK Then
'            Exit Sub
'        End If
'    End If
    Range("C2") = _
        "TopLeft:(" & rgSrc.Cells(1).Row & "," & rgSrc.Cells(1).Column & ")" & vbCr & _
        "Size:" & rgSrc.Rows.Count & "×" & rgSrc.Columns.Count
    
    Range("C3") = _
        "TopLeft:(" & rgDst.Cells(1).Row & "," & rgDst.Cells(1).Column & ")" & vbCr & _
        "Size:" & rgDst.Rows.Count & "×" & rgDst.Columns.Count
    
    '比較先の比較スタート地点の位置調整
    shDst.Activate
    cnt = rgDst.Count
    Set cell = Selection
    
    If cell.Column > rgDst.Cells(cnt).Column Then
        Set cell = shDst.Cells(cell.Row, rgDst.Cells(cnt).Column)
    End If
    If cell.Row > rgDst.Cells(cnt).Row Then
        Set cell = shDst.Cells(rgDst.Cells(cnt).Row, Cells.Column)
    End If
    cell.Select
    '比較終了点
    endA = cell.Address
    '比較開始
    Do
        Set cell = cell.Offset(0, 1)
        If cell.Column > rgDst.Cells(cnt).Column Then
            Set cell = shDst.Cells(cell.Row + 1, rgDst.Cells(1).Column)
            If cell.Row > rgDst.Cells(cnt).Row Then
                Set cell = rgDst.Cells(1)
            End If
        End If
        Debug.Print cell.Row & "," & cell.Column
        If cell.Text <> shSrc.Cells(cell.Row, cell.Column).Text Then
            shDst.Activate
            cell.Select
            'MsgBox "値違い"
            Exit Sub
        End If
        
    Loop While cell.Address <> endA
    MsgBox _
            shDst.Name & "と、" & shSrc.Name & vbCrLf & _
            "はピッタリ"
    Exit Sub
    
End Sub


