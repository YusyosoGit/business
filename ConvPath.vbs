Option Explicit
'BOX�p�X���ɕϊ�

Dim str
Const AppTitle = "�p�X���ϊ�"
Const BoxPath = "%homepath%"
Dim RE
Dim SS
Set RE = CreateObject("VBScript.RegExp")
Set SS = CreateObject("WScript.Shell")

str = GetClipboardText()
IF EditText(str) Then
    MsgBox "�ϊ��s�\�F" & vbCrLf & str, vbOK, AppTitle
    WScript.Quit 0
End If

Call PutInClipboardText(str)
MsgBox "�ϊ����܂����F" & str,vbOK, AppTitle


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
    Dim s
    'Dim mc, m

    EditText = 0
    s= str

    'P�h�� �� BOX�p�X
    RE.Pattern = "^P:."
    If RE.test(str) Then
        str = BoxPath & Mid(str, 3)
        Exit Function
    End If
    '���[�J���p�X �� BOX�p�X
    RE.Pattern = "^C:\\User\\\d{6}A00\u{4}\\..."
    if RE.test(str) Then
        str = BoxPath & Mid(str,10)
        Exit Function
    End If
    'BOX�p�X �� ���[�J���p�X
    RE.Pattern = "%homepath%..."
    If RE.test(str) Then
        str = "C:\User\112743A00BUJA\..." & Mid(str, 20)
        Exit Function
    End If

    '�ϊ��s�\
    EditText = 1
End Function