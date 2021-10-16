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
    Dim i As Integer
    Dim cell As Range
    Dim s As String, p As String
    Dim numbered As String
    Dim scenario As String
    
    Dim nMk As Integer
    Dim nRn As Integer
    Dim nMv As Integer

    Call DisNumber
    
    Set cell = Range("A1")
    i = 0
    
    Do While cell.Text <> ""
        s = Dir(cell.Text, vbDirectory)
        numbered = Right("00" & i, 3) & "_" & cell.Text
        If s = "" Then
            MkDir numbered
            nMk = nMk + 1
        Else
            '�ԍ���t�^����
            Name s As numbered
            nRn = nRn + 1
        End If
        
        'SetupData�ȊO�̃t�H���_�Ώ�
        If i <> 0 Then
            'scenario�t�H���_���쐬����
            scenario = numbered & "\scenario"
            s = Dir(scenario, vbDirectory)
            If s = "" Then
                MkDir scenario
            End If
        End If
        
        '�t�H���_���Ɩ��O�������t�@�C����������t�H���_�Ɉړ�
        s = Dir(cell.Text & ".*", vbNormal)
        Do While s <> ""
            Name s As numbered & "\scenario\" & s
            s = Dir()
        Loop
        i = i + 1
        Set cell = cell.Offset(1, 0)
    Loop
    
    MsgBox _
        nMk & "�t�H���_���쐬" & vbCrLf & _
        nRn & "�t�H���_���X�V" & vbCrLf & _
        nMv & "�̃t�@�C�����ړ����܂���"
    
End Sub

'����t�H���_����###_*�Ƃ������O�̃t�H���_��###_�����

Public Sub DisNumber()
    Dim i As Integer
    Dim s As String, path As String
    
    Dim exp As String
    exp = "*"
    
    path = ThisWorkbook.path
    'path = Left(path, InStrRev(path, "\"))
    's = Dir(path & "\" & exp)
    s = Dir(exp, vbDirectory)
    If s = "" Then Exit Sub
    
    Do While s <> ""
        If s Like "###_*" Then
        
            Debug.Print s
            Name _
                ThisWorkbook.path & "\" & s As _
                ThisWorkbook.path & "\" & Mid(s, 5)
        End If
        s = Dir()
    Loop
    
        
End Sub


