VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
Const Adr_SrcFilePath = "B2"
Const Adr_DstFilePath = "B3"
Const Adr_SrcShName = "C2"
Const Adr_DstShName = "C3"
Const Adr_SrcTblPsSz = "D2"
Const Adr_DstTblPsSz = "D3"


Const Adr_CompResult = "B4"
Dim vbWorkbookOpened

Public Sub CompSheets()
Attribute CompSheets.VB_ProcData.VB_Invoke_Func = "n\n14"
    Dim path
    Dim nmSrcBook As String
    Dim nmDstBook As String
    Dim nmSrcSh As String
    Dim nmDstSh As String
    Dim wb As Workbook
    Dim cnt
    Dim eAdr

    Dim cell As Range
    Dim rgDst As Range
    Dim rgSrc As Range
    Dim shDst As Worksheet
    Dim shSrc As Worksheet
    Dim bGetTableFromSheet As Boolean
    Dim ans
    
    vbWorkbookOpened = False
    
    nmSrcBook = Range(Adr_SrcFilePath).Text
    nmDstBook = Range(Adr_DstFilePath).Text
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
        
    bGetTableFromSheet = False
    If Range("A1") = "TRUE" Then
        bGetTableFromSheet = True
    End If
    If shSrc.Name <> Range(Adr_SrcShName).Text Then
        Range(Adr_SrcShName) = shSrc.Name
        bGetTableFromSheet = True
    End If
    If shDst.Name <> Range(Adr_DstShName).Text Then
        Range(Adr_DstShName) = shDst.Name
        bGetTableFromSheet = True
    End If
    
    '表の情報を記録から読み取る
    If Not bGetTableFromSheet Then
        Set rgSrc = ToRange(Range(Adr_SrcTblPsSz), shSrc)
        Set rgDst = ToRange(Range(Adr_DstTblPsSz), shDst)
        If rgSrc Is Nothing Or rgDst Is Nothing Then
            ans = MsgBox("記録から表の情報を読み取れませんでした。対象シートから取得しますか？", vbYesNo Or vbQuestion)
            If ans <> vbYes Then
                Exit Sub
            End If
            bGetTableFromSheet = True
        End If
    End If
        
    '表をシートから読み取る
    If bGetTableFromSheet Then
        Set rgSrc = shSrc.Range(shSrc.Cells(1), shSrc.Range("A1").SpecialCells(xlLastCell))
        Set rgDst = shDst.Range(shDst.Cells(1), shDst.Range("A1").SpecialCells(xlLastCell))
        Range(Adr_SrcTblPsSz) = ToPosSize(rgSrc)
        Range(Adr_DstTblPsSz) = ToPosSize(rgDst)
    End If
    
    '表の大きさが異なる場合の処理
    If Range(Adr_SrcTblPsSz).Text <> Range(Adr_DstTblPsSz).Text Then
        ans = MsgBox("比較する表の大きさが異なるので、２つのうちどちらかに合わせる必要があります。" & _
            "コピー元に合わせますか？" & vbCrLf & _
            "はい　：コピー元に合わせる" & vbCrLf & _
            "いいえ：コピー先に合わせる" & vbCrLf & _
            "キャンセル：中止", vbYesNoCancel Or vbQuestion, "シートの変更")
        If ans = vbCancel Then
            Exit Sub
        ElseIf ans = vbYes Then
            Range(Adr_SrcTblPsSz) = ToPosSize(rgSrc)
            Range(Adr_DstTblPsSz) = ToPosSize(rgSrc)
            Set rgDst = shDst.Range(shDst.Cells(rgSrc.Cells(1).Row, rgSrc.Cells(1).Column), _
                                    shDst.Cells(rgSrc.Cells(rgSrc.Cells.Count).Row, rgSrc.Cells(rgSrc.Cells.Count).Column))
        Else
            Range(Adr_SrcTblPsSz) = ToPosSize(rgDst)
            Range(Adr_DstTblPsSz) = ToPosSize(rgDst)
            Set rgSrc = shDst.Range(shSrc.Cells(rgDst.Cells(1).Row, rgDst.Cells(1).Column), _
                                    shSrc.Cells(rgDst.Cells(rgDst.Cells.Count).Row, rgDst.Cells(rgDst.Cells.Count).Column))
        End If
    End If
        
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
    
    nmDstBook = Range(Adr_DstFilePath).Text
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


' 範囲を(x, y),(x1, y1) に変換
Private Function ToPosSize(r As Range)
    ToPosSize = _
        "(" & r.Cells(1).Row & ", " & r.Cells(1).Column & ")," & _
        "(" & r.Rows.Count & ", " & r.Columns.Count & ")"
End Function


'座標を範囲に変換
Private Function ToRange(s As String, sh As Worksheet)
    Dim RE As New RegExp
    'Dim RE As Object
    'Set RE = CreateObject("VBScript.RegExp")
    Dim m As match
    Dim mc As MatchCollection
    
    
    RE.IgnoreCase = True
    RE.Global = True
    RE.Pattern = "[0-9]+" '"\([0-9]+, [0-9]+\)\-\([0-9]+, [0-9]+\)"
    Set mc = RE.Execute(s)
    If mc.Count <> 4 Then
        Set ToRange = Nothing
        Exit Function
    End If
    
    Set ToRange = sh.Range( _
                sh.Cells(Int(mc(0)), Int(mc(1))), _
                sh.Cells(Int(mc(2)), Int(mc(3))) _
            )
            
End Function
