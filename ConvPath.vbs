' ������ �p�X���ϊ� ������
' P�h����BOX�h���C�u�̃z�[���p�X
' ���[�J���p�X���z�[���p�X
' �z�[���p�X�����[�J���p�X

Option Explicit

Dim str
Const AppTitle = "�p�X���ϊ�"
Const HomePath = "%homepath%"
Const Label = "��n�ǃ\�t�g�E�F�A�J��_MW"
Dim RE
Dim SS
Dim s
Dim p1, p2

Set RE = CreateObject("VBScript.RegExp")
Set SS = CreateObject("WScript.Shell")

str = GetClipboardText()
'�����̕����i�s���j����菜��
str = Left(str, Len(str)-2)
s = str
IF ConvText(s) Then
    MsgBox "�ϊ��s�\�F" & vbCrLf & s, 0, AppTitle
    WScript.Quit 0
End If

Call PutInClipboardText(s)
p1 = InStr(str, Label)
p2 = InStr(s, Label)
MsgBox _ 
    Left(str, p1 + Len(Label)) & vbCrLf & _ 
    "��" & vbCrLf & _
    Left(s, p2 + Len(Label)) & vbCrLf & _
    Mid(s, p2 + Len(Label)+1), 0, AppTitle


' �N���b�v�{�[�h�ɂ���e�L�X�g��ǂݍ���
Public Function GetClipboardText()

    With SS
        With .Exec("powershell.exe -sta -Command Add-Type -Assembly System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()")
            .StdIn.Close
            GetClipboardText = .StdOut.ReadAll
        End With 
    End With
 
End Function

' �N���b�v�{�[�h�Ƀe�L�X�g���o��
Public Sub PutInClipboardText(ByVal str)

    With SS
        With .Exec("clip")
            Call .StdIn.Write(str)
        End With
    End With
End Sub

' ������ҏW
Public Function ConvText(ByRef str)

    'Dim t, u 'As Integer
    ConvText = 0
    RE.Pattern = "^"".+""$"
    If RE.test(str) Then
        str = Mid(str, 2, Len(str)-2)
    End If

    'P�h�� �� �z�[���p�X
    RE.Pattern = "^P:*"
    If RE.test(str) Then
        str = "C:" & HomePath & "\Box\��n�ǃ\�t�g�E�F�A�J��_MW" & Mid(str, 3)
        Exit Function
    End If
    RE.Pattern = "^C:"
    If Not RE.test(str) Then
        str = "C:" & str
    End If
    '�z�[���̃p�X �� ���[�J���p�X
    RE.Pattern = HomePath
    If RE.test(str) Then
        str = RE.Replace(str, "\User\112743A00BUJA")
        Exit Function
    End If
    '���[�J���p�X �� �z�[���p�X
    RE.Pattern = "\\User\\\d{6}A00\w{4}"
    If RE.test(str) Then
        str = RE.Replace(str, HomePath)
        Exit Function
    End If
    
    '�ϊ��s�\
    ConvText = 1
End Function