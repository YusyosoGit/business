Option Explicit

Dim s
Const Msg = "文字列をクリップボードにコピーしました。"

s = "お世話になります。" & vbCrLf & "お疲れ様です。" & vbCrLf & "【04月30日　金曜日】" & vbCrLf & "友松　創：09:00~17:00(7.0h)"
's = "お世話になります。" & ";" & "お疲れ様です。" '& vbCrLf & "【04月30日　金曜日】" & vbCrLf & "友松　創：09:00~17:00(7.0h)"

PutInClipboardText s
MsgBox Msg & vbCrLf & "文字列:" & s

Public Sub PutInClipboardText(ByVal str)
    '改行文字は出力されない
    'Dim cmd
    'cmd = "cmd /c ""echo " & str & "| clip"""
    'CreateObject("WScript.Shell").Run cmd, 0

    ' 改行文字も出力したい場合
    ' Dim ws
    ' Set ws = CreateObject("WScript.Shell")
    ' ws.Exec("clip").StdIn.Write str
    ' Set ws = Nothing

    'よりスマートに
    With CreateObject("WScript.Shell")
        With .Exec("clip")
            Call .StdIn.Write(str)
        End With
    End With
End Sub
