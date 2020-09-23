<div align="center">

## Barcode 3 of 9 generator


</div>

### Description

Create a 3 of 9 bar code without using a TTF, DLL or OCX.
 
### More Info
 
A String (sToCode) with the data you want to use for the bar code.

'A PictureBox (pPaintInto) object where the code will draw the bar code.

'A Label (pLabelInto) object where the code will put the human readable data.

You should have a Label control and a PictureBox control in your Form before running this code. This controls will be passed as parameters to the Sub.

'Example: Code3of9 "123-ABC", Picture1, Label1

'I see bar codes as a binary graphic since each "bar coded" character has a 16 pixels fixed width. Actually the standard says that the codes have 9 positions from which 3 of them are wide. Each position is either a bar or a space and between each character there is a narrow space. I added the space as part of the character.

'The sValidCodes string has all the valid characters bar codes coded into decimal numbers so every character uses 5 digits from the string. For example, the "*" character uses the last 5 digits from the string (35770) which would be:

'1000101110111010 in a binary way. Whenever there is a one, this code will draw a line in the propper position of the PictureBox control.

'I created this code to print the bar code that a friend uses in his frequent customer cards. He used to print them with several True Type Fonts and he was having a lot of trouble reading them with a (very) cheap bar code scanner. I developped this code originally for Excel but it has been changed to work perfectly in Visual Basic.

'If any one of you can test it in earlier versions of VB (4 or earlier) I would really appreciate your comments about how it works (if it works of course).

'And at the end but not less important please let me know what you think about it with your comments and/or score. I really appreciate that to know how I am doing with my development skills/knowledge.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Fernando Sola](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fernando-sola.md)
**Level**          |Intermediate
**User Rating**    |4.8 (105 globes from 22 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fernando-sola-barcode-3-of-9-generator__1-7179/archive/master.zip)





### Source Code

```
Public Sub Code3of9(sToCode As String, pPaintInto As PictureBox, pLabelInto As Label)
 Dim sValidChars As String
 Dim sValidCodes As String
 Dim lElevate As Integer
 Dim lCounter As Long
 Dim lWkValue As Long
 Dim PosX As Long
 Dim PosY1 As Long
 Dim PosY2 As Long
 Dim TPX As Long
 pPaintInto.Cls
 TPX = Screen.TwipsPerPixelX
 sValidChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%*"
 sValidCodes = "41914595664727860970419025962647338417105957" + _
 "84729059950476626106644590602984801043246599" + _
 "62476744460260046477586109044686603224803443" + _
 "91860130478424477058030365265828235758580903" + _
 "65863556658042365383495434978353624150635770"
 sToCode = UCase(IIf(Left(sToCode, 1) = "*", "", "*") + sToCode + IIf(Right(sToCode, 1) = "*", "", "*"))
 PosX = ((((pPaintInto.Width / TPX) - (Len(sToCode) * 16)) / 2) * TPX) - 1
 PosY1 = pPaintInto.Height * 0.2
 PosY2 = pPaintInto.Height * 0.8
 If PosX < 0 Then
 MsgBox "The length of the code exceeds control limits.", vbExclamation, "Large string"
 GoTo End_Code
 End If
 On Error Resume Next
 For lCounter = 1 To Len(sToCode)
'Here is where the number is fetched from the sValidCodes string. It will get only 5 digits.
 lWkValue = Val(Mid(sValidCodes, ((InStr(1, sValidChars, Mid(sToCode, lCounter, 1)) - 1) * 5) + 1, 5))
 lWkValue = IIf(lWkValue = 0, 36538, lWkValue)
 For lElevate = 15 To 0 Step -1
 'It evaluates the binary number to see if it has to draw a line.
 If lWkValue >= 2 ^ lElevate Then
 pPaintInto.Line (PosX, PosY1)-(PosX, PosY2)
 lWkValue = lWkValue - (2 ^ lElevate)
 End If
 PosX = PosX + TPX
 Next
 Next
 pLabelInto.Caption = Mid(sToCode, 2, Len(sToCode) - 2)
End_Code:
End Sub
```

