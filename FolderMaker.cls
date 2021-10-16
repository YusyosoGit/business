VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'３桁数字付きのフォルダを作成
'同一フォルダ内の同じ名のファイルを出来たフォルダに移動

Public Sub FolderMaker()
    Dim i As Integer
    Dim cell As Range
    Dim s As String, p As String
    Dim numbered As String
    Dim scenario As String
    
    Dim nMk As Integer
    Dim nRn As Integer
    Dim nMv As Integer

    'カレントディレクトリの変更
    ChDir ThisWorkbook.path

    '３桁の番号を取る
    Call DisNumber
    
    Set cell = Range("A1")
    i = 0
    
    Do While cell.Text <> ""
        s = Dir(cell.Text, vbDirectory)
        numbered = Right("00" & i, 3) & "_" & cell.Text
        If s = "" Then
            MkDir numbered
            nMk = nMk + 1
        Else
            '番号を付与する
            Name s As numbered
            nRn = nRn + 1
        End If
        
        'SetupData以外のフォルダ対象
        If i <> 0 Then
            'scenarioフォルダを作成する
            scenario = numbered & "\scenario"
            s = Dir(scenario, vbDirectory)
            If s = "" Then
                MkDir scenario
            End If
        End If
        
        'フォルダ名と名前が同じファイルを作ったフォルダに移動
        s = Dir(cell.Text & ".*", vbNormal)
        Do While s <> ""
            Name s As numbered & "\scenario\" & s
            s = Dir()
        Loop
        i = i + 1
        Set cell = cell.Offset(1, 0)
    Loop
    
    MsgBox _
        nMk & "つフォルダを作成" & vbCrLf & _
        nRn & "つフォルダを更新" & vbCrLf & _
        nMv & "つのファイルを移動しました"
    
End Sub

'同一フォルダ内の###_*という名前のフォルダの###_を取る

Public Sub DisNumber()
    Dim i As Integer
    Dim s As String, path As String
    
    Dim exp As String
    exp = "*"
    
    'path = Left(path, InStrRev(path, "\"))
    's = Dir(path & "\" & exp)
    s = Dir(exp, vbDirectory)
    If s = "" Then Exit Sub
    
    Do While s <> ""
        If s Like "###_*" Then
        
            Debug.Print s
            Name _
                ThisWorkbook.path & "\" & s As _
                ThisWorkbook.path & "\" & Mid(s, 5)
        End If
        s = Dir()
    Loop
    
        
End Sub


