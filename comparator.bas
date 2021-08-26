' ファイルAとファイルBの表示中のシートを比較するプログラム
' 単純な見たままの値の比較で、書式や数式は含みません

Option Explicit

'このコードを記載しているシートのセルに次の値を設定しておく
'B2:比較元のブック名
'B3:比較先のブック名
'比較対象の両方のブックの該当シートをあらかじめ開いておくこと。

'プログラム実行後
'２つの表の値に差異のあるセルに、カーソルを置いて処理をストップ。結果表示セルにそのアドレス表示
'値が等しいならメッセージを表示。結果表示セルに「差異なし」表示
Const Adr_SrcFileInfo = "B2"
Const Adr_DstFileInfo = "B3"
Const Adr_CompResult = "B4"
Dim vbWorkbookOpened

Public Sub CompSheets()
    Dim path
    Dim nmSrcBook As String
    Dim nmDstBook As String
    Dim wb As Workbook
    Dim cnt
    Dim eAdr

    Dim cell As Range
    Dim rgDst As Range
    Dim rgSrc As Range
    Dim shDst As Worksheet
    Dim shSrc As Worksheet
    Dim match As Boolean
    Dim ans
    
    vbWorkbookOpened = False
    
    nmSrcBook = Range(Adr_SrcFileInfo).Text
    nmDstBook = Range(Adr_DstFileInfo).Text
    'Debug.Print Workbooks(Range("B3").Text).Worksheets.Count
    
    '処理の対象はこのコードを実行しているワークブックと同じパスのエクセルファイル
    'path = hisWorkbook.path & "\"
    Set wb = OpenBook(nmSrcBook, True)
    If wb Is Nothing Then
        MsgBox "コピー元ファイルなし", vbExclamation
        Exit Sub
    Else
        Set shSrc = wb.ActiveSheet
    End If
    
    Set wb = OpenBook(nmDstBook, False)
    If wb Is Nothing Then
        MsgBox "コピー先ファイルなし", vbExclamation
        Exit Sub
    Else
        Set shDst = wb.ActiveSheet
    End If
    
    If vbWorkbookOpened = True Then
        MsgBox "シートを選択してもう一度実行してください。", vbInformation
        Exit Sub
    End If
        
    Set rgSrc = shSrc.Range(shSrc.Cells(1, 1), shSrc.Range("A1").SpecialCells(xlLastCell))
    Set rgDst = shDst.Range(shDst.Cells(1, 1), shDst.Range("A1").SpecialCells(xlLastCell))
    
    Debug.Print rgSrc.Rows.Count
    Debug.Print rgDst.Rows.Count
    Debug.Print rgSrc.Columns.Count
    Debug.Print rgDst.Columns.Count
    
    Range("C2") = _
        "ShName:" & shSrc.Name & vbCr & _
        "TopLeft:(" & rgSrc.Cells(1).Row & "," & rgSrc.Cells(1).Column & ")" & vbCr & _
        "Size:" & rgSrc.Rows.Count & "×" & rgSrc.Columns.Count
    
    Range("C3") = _
        "ShName:" & shDst.Name & vbCr & _
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
    eAdr = cell.Address
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
            Range(Adr_CompResult) = cell.Address()
            shDst.Activate
            cell.Select
            'MsgBox "値違い"
            Exit Sub
        End If
        
    Loop While cell.Address <> eAdr
    Range(Adr_CompResult) = "差異なし"
    MsgBox _
            shDst.Name & "と、" & shSrc.Name & vbCrLf & _
            "はピッタリ", _
            vbInformation
            
    Exit Sub
    
End Sub



'黄色←⇒橙色
Public Sub InitSheetColorLTE()
    Dim path
    Dim nmDstBook As String
    Dim nm As String
    Dim cnt

    Dim shDst As Worksheet
    Dim i, iAns
    Dim x
    
    nmDstBook = Range(Adr_DstFileInfo).Text
    'Debug.Print Workbooks(Range("B3").Text).Worksheets.Count
    
    nm = Dir(path & nmDstBook)
    If nm <> "" Then
        Set shDst = Workbooks(nm).ActiveSheet
    Else
        MsgBox "コピー先ファイルなし"
        Exit Sub
    End If
    
    iAns = MsgBox("色付きのタブの色を一律にリセットします", vbOKCancel)
    If iAns <> vbOK Then Exit Sub
    For i = 3 To Sheets.Count
        With Worksheets(i).Tab
            x = .ColorIndex
            If .Color = rgbOrange Then
                .Color = rgbYellow
            End If
        End With
    Next i
End Sub


Private Function OpenBook(path As String, bReadOnly As Boolean)
    Dim wb As Workbook
    Dim nm As String
    Dim found
    Set OpenBook = Nothing
    
    nm = Dir(path)
    If nm = "" Then
        Exit Function
    End If
    
    For Each wb In Workbooks
        If wb.Name = nm Then
            found = True
            Exit For
        End If
    Next
    
    If found = True Then
        Set OpenBook = Workbooks(nm)
        Exit Function
    End If
    
    
    Set OpenBook = Workbooks.Open(Filename:=path, ReadOnly:=bReadOnly)
    vbWorkbookOpened = True
End Function
