VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'���̃R�[�h���L�ڂ��Ă���V�[�g�̃Z���Ɏ��̒l��ݒ肵�Ă���
'B2:��r���̃u�b�N��
'B3:��r��̃u�b�N��
'��r�Ώۂ̗����̃u�b�N�̊Y���V�[�g�����炩���ߊJ���Ă������ƁB

'�v���O�������s��
'�l�ɈႢ����������A���̃Z����I�����ď������X�g�b�v
'�l���������Ȃ烁�b�Z�[�W��\�����ďI������

Public Sub CompSheets()
    'Dim path
    Dim nmSrcBook As String
    Dim nmDstBook As String
    Dim nmSrcSheet As String
    Dim nmDstSheet As String

    Dim cell As Range
    Dim rgDst As Range
    Dim rgSrc As Range
    Dim shDst As Worksheet
    Dim shSrc As Worksheet
    Dim match As Boolean
    
    'path = Range("B1").Text
    'If Right(path, 1) <> "\" Then path = path & "\"
    
    nmSrcBook = Range("B2").Text
    nmDstBook = Range("B3").Text
    'Debug.Print Workbooks(Range("B3").Text).Worksheets.Count
    
    '�����̑Ώۂ͂��̃R�[�h�����s���Ă��郏�[�N�u�b�N�Ɠ����p�X�̃G�N�Z���t�@�C��
    Set shDst = Workbooks(nmDstBook).ActiveSheet
    Set shSrc = Workbooks(nmSrcBook).ActiveSheet
    '�͈͎w��̑Ώۂ�3�������ׂĂɁA�V�[�g�����܂߂Ȃ��Ƃ����Ȃ�
    Set rgSrc = shSrc.Range(shSrc.Cells(1, 1), shSrc.Range("A1").SpecialCells(xlLastCell))
    Set rgDst = shDst.Range(shDst.Cells(1, 1), shDst.Range("A1").SpecialCells(xlLastCell))
    
    Debug.Print rgSrc.Rows.Count
    Debug.Print rgDst.Rows.Count
    Debug.Print rgSrc.Columns.Count
    Debug.Print rgDst.Columns.Count
    
    If rgDst.Rows.Count <> rgSrc.Rows.Count Or _
        rgDst.Columns.Count <> rgSrc.Columns.Count Then
        ' https://daitaideit.com/vba-range/
        MsgBox "�T�C�Y�Ⴂ"
        Exit Sub
    End If
    
    match = True

    For Each cell In rgDst
        Debug.Print cell.Row & "," & cell.Column
        If cell.Value <> rgSrc(cell.Row, cell.Column).Value Then
            shDst.Activate
            cell.Select
            'MsgBox "�l�Ⴂ"
            Exit Sub
        End If
    Next
    MsgBox _
            shDst.Name & "�ƁA" & shSrc.Name & vbCrLf & _
            "�̓��e����v���܂����B"
End Sub