Option Explicit

Dim str
str = GetClipboardText()
IF EditText(str) Then
    MsgBox "�N���b�v�{�[�h�̕�����͕s���ł��B"
    WScript.Quit 0
End If

Call PutInClipboardText(str)
MsgBox "�N���b�v�{�[�h�ɏo�͂��܂����B" & vbCrLf & str


' �N���b�v�{�[�h�ɂ���e�L�X�g��ǂݍ���
Public Function GetClipboardText()

    With CreateObject("WScript.Shell")
        With .Exec("powershell.exe -sta -Command Add-Type -Assembly System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()")
            .StdIn.Close
            GetClipboardText = .StdOut.ReadAll
        End With 
    End With
 
End Function

' �N���b�v�{�[�h�Ƀe�L�X�g���o��
Public Sub PutInClipboardText(ByVal str)

    With CreateObject("WScript.Shell")
        With .Exec("clip")
            Call .StdIn.Write(str)
        End With
    End With
End Sub

' ������ҏW
Public Function EditText(ByRef str)

    Dim t, u 'As Integer
    Dim s, p

    EditText = 0
    s= str

    t = InStr(s, "�y")
    If t = 0 Then '���x����Goto�͎g���Ȃ��i�G���[�j
        EditText = 1
        Exit Function
    End If
    s = Mid(s, t)

    't = InStr(s, "�z")
    t = InStr(s, vbCrLf)
    u = InStr(s, "�F")
    if t = 0 Or u = 0 Then 
        EditText = 1
        Exit Function
    End If
    s = Left(s, t) & Mid(s, u+1)

    str = s
End Function