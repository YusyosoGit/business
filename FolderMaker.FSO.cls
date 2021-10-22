VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'�R�������t���̃t�H���_���쐬
'����t�H���_���̓������̃t�@�C�����o�����t�H���_�Ɉړ�


Public Sub FolderMaker()
    Dim FSO As New FileSystemObject
    Dim Fo As Folder
    Dim ThisFo As Folder
    Dim Fi As File
    
    Dim RE As New RegExp
    RE.Pattern = "^[0-9]{3}_."
    RE.Global = False
    Dim mc As MatchCollection
    Dim m As Match
    
    Dim i As Integer
    Dim cell As Range
    Dim s As String, p As String
    Dim numbered As String
    Dim scenario As String
    Dim moved As String
    
    Dim found As Boolean
    
    Dim nMk As Integer
    Dim nRn As Integer
    Dim nMv As Integer

    '�J�����g�f�B���N�g���̕ύX
    ChDir ThisWorkbook.path

    '�R���̔ԍ������
    Set ThisFo = FSO.GetFolder(ThisWorkbook.path)
    For Each Fo In ThisFo.SubFolders
'        Set mc = RE.Execute(Fo.Name)
        If RE.Test(Fo.Name) Then
            Fo.Name = Mid(Fo.Name, 5)
        End If
    Next
    
    Set cell = Range("A1")
    i = 0
    
    '�V�[�g�̃t�H���_�������A�ė��p�����肷��
    Do While cell.Text <> ""
        numbered = Right("00" & i, 3) & "_" & cell.Text
        If FSO.FolderExists(".\" & cell.Text) Then
            FSO.MoveFolder ".\" & cell.Text, ".\" & numbered
            nRn = nRn + 1
        Else
            FSO.CreateFolder numbered
            nMk = nMk + 1
        End If
        
        'SetupData�ȊO�̃t�H���_�Ώ�
        If i <> 0 Then
            'scenario�t�H���_���쐬����
            scenario = numbered & "\scenario"
            If Not FSO.FolderExists(".\" & scenario) Then
                FSO.CreateFolder scenario
            End If
        
            '�t�H���_���Ɩ��O�������t�@�C����������t�H���_�Ɉړ�
            For Each Fi In ThisFo.Files
                If cell.Text = Left(Fi.Name, Len(cell.Text)) Then
                    '�g���q�𒲂ׂ�
                    Select Case Mid(Fi.Name, Len(cell.Text) + 1)
                    Case ".xlsm", ".pcap", ".dat"
                        moved = scenario & "\" & Fi.Name
                        If FSO.FileExists(".\" & moved) Then
                            FSO.DeleteFile ".\" & moved
                        End If
                        FSO.MoveFile ".\" & Fi.Name, ".\" & moved
                        nMv = nMv + 1
                    End Select
                End If
            Next
        End If
        i = i + 1
        Set cell = cell.Offset(1, 0)
    Loop
    
    MsgBox _
        nMk & "�t�H���_���쐬" & vbCrLf & _
        nRn & "�t�H���_���X�V" & vbCrLf & _
        nMv & "�̃t�@�C�����ړ����܂���"
    
End Sub
