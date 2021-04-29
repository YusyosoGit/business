Option Explicit

Dim str
str = GetClipboardText()
IF EditText(str) Then
    MsgBox "クリップボードの文字列は不正です。"
    WScript.Quit 0
End If

Call PutInClipboardText(str)
MsgBox "クリップボードに出力しました。" & vbCrLf & str


' クリップボードにあるテキストを読み込む
Public Function GetClipboardText()

    With CreateObject("WScript.Shell")
        With .Exec("powershell.exe -sta -Command Add-Type -Assembly System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()")
            .StdIn.Close
            GetClipboardText = .StdOut.ReadAll
        End With 
    End With
 
End Function

' クリップボードにテキストを出力
Public Sub PutInClipboardText(ByVal str)

    With CreateObject("WScript.Shell")
        With .Exec("clip")
            Call .StdIn.Write(str)
        End With
    End With
End Sub

' 文字列編集
Public Function EditText(ByRef str)

    Dim t, u 'As Integer
    Dim s, p

    EditText = 0
    s= str

    t = InStr(s, "【")
    If t = 0 Then 'ラベル＆Gotoは使えない（エラー）
        EditText = 1
        Exit Function
    End If
    s = Mid(s, t)

    't = InStr(s, "】")
    t = InStr(s, vbCrLf)
    u = InStr(s, "：")
    if t = 0 Or u = 0 Then 
        EditText = 1
        Exit Function
    End If
    s = Left(s, t) & Mid(s, u+1)

    str = s
End Function