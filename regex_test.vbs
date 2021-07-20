Option Explicit

Dim str
Dim title
str = GetClipboardText()
IF EditText(str, title) Then
    MsgBox "クリップボードの文字列は不正です。"
    WScript.Quit 0
End If

MsgBox title

'Call PutInClipboardText(str)
'MsgBox "クリップボードに出力しました。" & vbCrLf & str


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
Public Function EditText(ByRef str, ByRef title)
    Dim RE
    Dim m, mc
    EditText = 0
    
    Set RE = CreateObject("VBScript.RegExp")
    RE.Global = True
    RE.IgnoreCase = False
    RE.Pattern = "\([0-9]+\)"
    
    Set mc = RE.Execute(str)
    If mc.Count = 0 Then
        EditText = 1
        Exit Function
    End If

    Set m = mc(0)
    title = Left(str, m.FirstIndex + m.Length)

End Function