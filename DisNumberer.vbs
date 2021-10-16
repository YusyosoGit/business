'*
'* �T�u�t�H���_���܂߂ăt�@�C�����E�t�H���_������擪�ԍ����������܂�
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

'���̃t�@�C����u���Ă���f�B���N�g�����Ώ�
Path = FSO.GetFolder(".").Path
ans = MsgBox( _
    "�t�H���_���E�t�@�C��������擪�̔ԍ����������܂��B" & vbCrLf & _
    "�����t�H���_�F" & vbCrLf & _
    Path & vbCrLf, vbOKCancel)
If ans <> vbOK Then
    WScript.Quit 0
End If

'�{����
Call RenameFilesAndFolders(Path)

If NumFo = 0 And NumFi = 0 Then
    MsgBox "�����B" & vbCrLf & "�t�@�C�����E�t�H���_���ɕύX�Ȃ��B"
    WScript.Quit 0
End If

MsgBox "�����B" & vbCrLf &_
    "�ύX�t�H���_���F" & NumFo & vbCrLf & _
    "�ύX�t�@�C�����F" & NumFi

Public Sub RenameFilesAndFolders(targetDir)
    Dim Fo, Fi
    Dim ThisFo
    Dim mc, m
    'Dim FSO
    'Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
 
    '�t�H���_���ύX
    Set ThisFo = FSO.GetFolder(targetDir)
    Set mc = RE.Execute(ThisFo.Name)
    If mc.Count = 1 Then
        ThisFo.Name = Mid(ThisFo.Name, 5)
        NumFo = NumFo + 1
    End If

    '�t�H���_������
    For Each Fo In ThisFo.SubFolders
        Call RenameFilesAndFolders(Fo.Path)
    Next

    '�t�@�C�����ύX
    For Each Fi In ThisFo.Files
        Set mc = RE.Execute(Fi.Name)
        If mc.Count = 1 Then
            Fi.Name = Mid(Fi.Name, 5)
            NumFi = NumFi + 1
        End If
    Next

End Sub        
