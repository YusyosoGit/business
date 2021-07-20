'Option Explicit

Dim str
Dim appOL
Set appOL = CreateObject("Outlook.Application")
Dim mail 
'Set mail = appOL.CreateItem("Outlook.olMailItem")
Set mail = appOL.CreateItem(0) 'olMailItem

With mail
    .To = "tomomatsu-it@mbr.nifty.com"
    .Subject = "‚ ‚¢‚³‚Â"
    .Body = "‚±‚ñ‚Î‚ñ‚Í"
    .BodyFormat = 1 'olFormatPlain
    '.Display True
End With

Dim bAns
bAns = MsgBox("ƒ[ƒ‹‚ğ‘—‚è‚Ü‚·‚©H", vbYesNo)
If (bAns <> vbYes) Then
    WScript.Quit
End If
mail.Send

Set mail = Nothing
Set appOL = Nothing
