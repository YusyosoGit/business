Option Explicit

Dim s
Const Msg = "��������N���b�v�{�[�h�ɃR�s�[���܂����B"

s = "�����b�ɂȂ�܂��B" & vbCrLf & "�����l�ł��B" & vbCrLf & "�y04��30���@���j���z" & vbCrLf & "�F���@�n�F09:00~17:00(7.0h)"
's = "�����b�ɂȂ�܂��B" & ";" & "�����l�ł��B" '& vbCrLf & "�y04��30���@���j���z" & vbCrLf & "�F���@�n�F09:00~17:00(7.0h)"

PutInClipboardText s
MsgBox Msg & vbCrLf & "������:" & s

Public Sub PutInClipboardText(ByVal str)
    '���s�����͏o�͂���Ȃ�
    'Dim cmd
    'cmd = "cmd /c ""echo " & str & "| clip"""
    'CreateObject("WScript.Shell").Run cmd, 0

    ' ���s�������o�͂������ꍇ
    ' Dim ws
    ' Set ws = CreateObject("WScript.Shell")
    ' ws.Exec("clip").StdIn.Write str
    ' Set ws = Nothing

    '���X�}�[�g��
    With CreateObject("WScript.Shell")
        With .Exec("clip")
            Call .StdIn.Write(str)
        End With
    End With
End Sub
