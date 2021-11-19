Option Explicit
'BOXパス名に変換

Dim str
Const AppTitle = "パス名変換"
Const BoxPath = "%homepath%"
Dim RE
Dim SS
Set RE = CreateObject("VBScript.RegExp")
Set SS = CreateObject("WScript.Shell")

str = GetClipboardText()
IF EditText(str) Then
    MsgBox "変換不可能：" & vbCrLf & str, vbOK, AppTitle
    WScript.Quit 0
End If

Call PutInClipboardText(str)
MsgBox "変換しました：" & str,vbOK, AppTitle


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
    Dim s
    'Dim mc, m

    EditText = 0
    s= str

    'Pドラ → BOXパス
    RE.Pattern = "^P:."
    If RE.test(str) Then
        str = BoxPath & Mid(str, 3)
        Exit Function
    End If
    'ローカルパス → BOXパス
    RE.Pattern = "^C:\\User\\\d{6}A00\u{4}\\..."
    if RE.test(str) Then
        str = BoxPath & Mid(str,10)
        Exit Function
    End If
    'BOXパス → ローカルパス
    RE.Pattern = "%homepath%..."
    If RE.test(str) Then
        str = "C:\User\112743A00BUJA\..." & Mid(str, 20)
        Exit Function
    End If

    '変換不可能
    EditText = 1
End Function