VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'�t�H���_���ꗗ�\��
'�g�p���@�͂��̃t�@�C�����ꗗ��m�肽���t�H���_�ɒu���āA�{�^���N���b�N

Dim row As Integer
Dim RenFo As Integer
Dim RenFi As Integer

Public Sub DisNumberer()
'   �Q�Ɛݒ� Microsoft Scripting Runtime
    Dim FSO As New FileSystemObject
    Dim Fo As Folder
    
    RenFo = 0
    RenFi = 0
    
    Set Fo = FSO.GetFolder(ThisWorkbook.Path)
    row = 1
    Call SetFolderInfo(Fo.Path, 1)
    
    MsgBox _
        RenFi & "�̃t�@�C������ύX" & vbCrLf & _
        RenFo & "�̃t�H���_����ύX"
        
End Sub



Public Sub SetFolderInfo(targetDir As String, col As Integer)
    Dim FSO As New FileSystemObject
    Dim ThisFO As Folder
    Dim Fo As Folder
    Dim Fi As File
    Const Fmt = "###_*"
    
    
    Set ThisFO = FSO.GetFolder(targetDir)
    If ThisFO.name Like Fmt Then
        '�t�H���_���ύX
        ThisFO.name = Mid(ThisFO.name, 5)
        RenFo = RenFo + 1
    End If
    For Each Fo In ThisFO.SubFolders
        row = row + 1
        Call SetFolderInfo(Fo.Path, col + 1)
    Next
    For Each Fi In ThisFO.Files
        If Fi.name Like Fmt Then
            '�t�@�C�����ύX
            Fi.name = Mid(ThisFO.name, 5)
            RenFi = RenFi + 1
        End If
    Next
        
End Sub
