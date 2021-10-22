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
    Dim FSO As New FileSystemObject
    Dim Fo As Folder
    Dim ThisFo As Folder
    Dim Fi As File
    
    Dim RE As New RegExp
    RE.Pattern = "^[0-9]{3}_."
    RE.Global = False
    Dim mc As MatchCollection
    Dim m As Match
    
    Dim i As Integer
    Dim cell As Range
    Dim s As String, p As String
    Dim numbered As String
    Dim scenario As String
    Dim moved As String
    
    Dim found As Boolean
    
    Dim nMk As Integer
    Dim nRn As Integer
    Dim nMv As Integer

    'カレントディレクトリの変更
    ChDir ThisWorkbook.path

    '３桁の番号を取る
    Set ThisFo = FSO.GetFolder(ThisWorkbook.path)
    For Each Fo In ThisFo.SubFolders
'        Set mc = RE.Execute(Fo.Name)
        If RE.Test(Fo.Name) Then
            Fo.Name = Mid(Fo.Name, 5)
        End If
    Next
    
    Set cell = Range("A1")
    i = 0
    
    'シートのフォルダ作ったり、再利用したりする
    Do While cell.Text <> ""
        numbered = Right("00" & i, 3) & "_" & cell.Text
        If FSO.FolderExists(".\" & cell.Text) Then
            FSO.MoveFolder ".\" & cell.Text, ".\" & numbered
            nRn = nRn + 1
        Else
            FSO.CreateFolder numbered
            nMk = nMk + 1
        End If
        
        'SetupData以外のフォルダ対象
        If i <> 0 Then
            'scenarioフォルダを作成する
            scenario = numbered & "\scenario"
            If Not FSO.FolderExists(".\" & scenario) Then
                FSO.CreateFolder scenario
            End If
        
            'フォルダ名と名前が同じファイルを作ったフォルダに移動
            For Each Fi In ThisFo.Files
                If cell.Text = Left(Fi.Name, Len(cell.Text)) Then
                    '拡張子を調べる
                    Select Case Mid(Fi.Name, Len(cell.Text) + 1)
                    Case ".xlsm", ".pcap", ".dat"
                        moved = scenario & "\" & Fi.Name
                        If FSO.FileExists(".\" & moved) Then
                            FSO.DeleteFile ".\" & moved
                        End If
                        FSO.MoveFile ".\" & Fi.Name, ".\" & moved
                        nMv = nMv + 1
                    End Select
                End If
            Next
        End If
        i = i + 1
        Set cell = cell.Offset(1, 0)
    Loop
    
    MsgBox _
        nMk & "つフォルダを作成" & vbCrLf & _
        nRn & "つフォルダを更新" & vbCrLf & _
        nMv & "つのファイルを移動しました"
    
End Sub

