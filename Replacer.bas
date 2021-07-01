Attribute VB_Name = "Module1"
Option Explicit

Public Sub Macro_Replace()
    Call ReplaceRegEx("[a-zA-Z]*:", "Hoverer::")
End Sub


'http://www.eurus.dti.ne.jp/~yoneyama/Excel/vba/vba_regexp.html
'�V�[�g���̃Z���̐��K�\���ɍ��v����L�[���[�h��u������
Public Function ReplaceRegEx(strKeyword As String, strAfter As String)
    Dim RE
    Dim r As Range, endP As String
    ReplaceRegEx = 0
    
    Set RE = CreateObject("VBScript.RegExp")
    RE.Global = True
    RE.IgnoreCase = False
    RE.Pattern = strKeyword
    
    Set r = Cells.Find("*")
    If r Is Nothing Then GoTo EndFunction
    endP = r.Address
    Do
        If RE.Test(r.Text) Then
            r = RE.Replace(r.Text, strAfter)
        End If
        Set r = Cells.FindNext(r)
    Loop Until r.Address = endP
    
EndFunction:
    '�I������

End Function

