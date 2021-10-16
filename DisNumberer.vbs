'*
'* サブフォルダを含めてファイル名・フォルダ名から先頭番号を除去します
'*
Option Explicit
Dim FSO 
Dim RE
Dim Path

Dim NumFo
Dim NumFi
Dim ans

Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
Set RE = CreateObject("VBScript.RegExp")
RE.Pattern = "^[0-9]{3}_."

NumFo = 0
NumFi = 0

'このファイルを置いているディレクトリが対象
Path = FSO.GetFolder(".").Path
ans = MsgBox( _
    "フォルダ名・ファイル名から先頭の番号を除去します。" & vbCrLf & _
    "検索フォルダ：" & vbCrLf & _
    Path & vbCrLf, vbOKCancel)
If ans <> vbOK Then
    WScript.Quit 0
End If

'本処理
Call RenameFilesAndFolders(Path)

If NumFo = 0 And NumFi = 0 Then
    MsgBox "完了。" & vbCrLf & "ファイル名・フォルダ名に変更なし。"
    WScript.Quit 0
End If

MsgBox "完了。" & vbCrLf &_
    "変更フォルダ名：" & NumFo & vbCrLf & _
    "変更ファイル名：" & NumFi

Public Sub RenameFilesAndFolders(targetDir)
    Dim Fo, Fi
    Dim ThisFo
    Dim mc, m
    'Dim FSO
    'Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
 
    'フォルダ名変更
    Set ThisFo = FSO.GetFolder(targetDir)
    Set mc = RE.Execute(ThisFo.Name)
    If mc.Count = 1 Then
        ThisFo.Name = Mid(ThisFo.Name, 5)
        NumFo = NumFo + 1
    End If

    'フォルダ内検索
    For Each Fo In ThisFo.SubFolders
        Call RenameFilesAndFolders(Fo.Path)
    Next

    'ファイル名変更
    For Each Fi In ThisFo.Files
        Set mc = RE.Execute(Fi.Name)
        If mc.Count = 1 Then
            Fi.Name = Mid(Fi.Name, 5)
            NumFi = NumFi + 1
        End If
    Next

End Sub        
