' ■■■ パス名変換 ■■■
' Pドラ→BOXドライブのホームパス
' ローカルパス→ホームパス
' ホームパス→ローカルパス

Option Explicit

Dim str
Const AppTitle = "パス名変換"
Const HomePath = "%homepath%"
Const Label = "基地局ソフトウェア開発_MW"
Dim RE
Dim SS
Dim s
Dim p1, p2

Set RE = CreateObject("VBScript.RegExp")
Set SS = CreateObject("WScript.Shell")

str = GetClipboardText()
'文末の文字（不明）を取り除く
str = Left(str, Len(str)-2)
s = str
IF ConvText(s) Then
    MsgBox "変換不可能：" & vbCrLf & s, 0, AppTitle
    WScript.Quit 0
End If

Call PutInClipboardText(s)
p1 = InStr(str, Label)
p2 = InStr(s, Label)
MsgBox _ 
    Left(str, p1 + Len(Label)) & vbCrLf & _ 
    "↓" & vbCrLf & _
    Left(s, p2 + Len(Label)) & vbCrLf & _
    Mid(s, p2 + Len(Label)+1), 0, AppTitle


' クリップボードにあるテキストを読み込む
Public Function GetClipboardText()

    With SS
        With .Exec("powershell.exe -sta -Command Add-Type -Assembly System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()")
            .StdIn.Close
            GetClipboardText = .StdOut.ReadAll
        End With 
    End With
 
End Function

' クリップボードにテキストを出力
Public Sub PutInClipboardText(ByVal str)

    With SS
        With .Exec("clip")
            Call .StdIn.Write(str)
        End With
    End With
End Sub

' 文字列編集
Public Function ConvText(ByRef str)

    'Dim t, u 'As Integer
    ConvText = 0
    RE.Pattern = "^"".+""$"
    If RE.test(str) Then
        str = Mid(str, 2, Len(str)-2)
    End If

    'Pドラ → ホームパス
    RE.Pattern = "^P:*"
    If RE.test(str) Then
        str = "C:" & HomePath & "\Box\基地局ソフトウェア開発_MW" & Mid(str, 3)
        Exit Function
    End If
    RE.Pattern = "^C:"
    If Not RE.test(str) Then
        str = "C:" & str
    End If
    'ホームのパス → ローカルパス
    RE.Pattern = HomePath
    If RE.test(str) Then
        str = RE.Replace(str, "\User\112743A00BUJA")
        Exit Function
    End If
    'ローカルパス → ホームパス
    RE.Pattern = "\\User\\\d{6}A00\w{4}"
    If RE.test(str) Then
        str = RE.Replace(str, HomePath)
        Exit Function
    End If
    
    '変換不可能
    ConvText = 1
End Function