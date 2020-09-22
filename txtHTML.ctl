VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl txtHTML 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ScaleHeight     =   1215
   ScaleWidth      =   1215
   ToolboxBitmap   =   "txtHTML.ctx":0000
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2143
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   1e7
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"txtHTML.ctx":0314
   End
End
Attribute VB_Name = "txtHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim apppath As String
Dim starttime As Date
Dim tmpchr As String * 1
Dim tmpint As Long
Dim varColorText, varColorTag, varColorProp, varColorPropVal, varColorComment As OLE_COLOR

Public Event RightClick()
Public Event LeftClick()
Public Event Change()
Public Event DropFile(FileName As String)

Const EM_CANUNDO = &HC6
Const EM_UNDO = &HC7

Function CanUndo() As Boolean
  CanUndo = SendMessage(rtf1.hwnd, EM_CANUNDO, 0&, 0&)
End Function

Function Undo() As Boolean
  SendMessage rtf1.hwnd, EM_UNDO, 0&, 0&
End Function

Function Cut()
  With Clipboard
     .Clear
     .SetText rtf1.SelText, vbCFText
     rtf1.SelText = ""
  End With
End Function

Function Copy()
  With Clipboard
     .Clear
     .SetText rtf1.SelText, vbCFText
  End With
End Function

Function Paste()
  rtf1.SelText = Clipboard.GetText(vbCFText)
  rtf1.SelStart = 0
  rtf1.SelLength = Len(rtf1.Text)
  ColorHTML
  rtf1.SelStart = 1
End Function

Function SelectAll()
  rtf1.SelStart = 0
  rtf1.SelLength = Len(rtf1.Text)
  rtf1.SetFocus
End Function

Function Delete()
  rtf1.SelText = ""
End Function

Function fcnGetRTFColor(ByVal Color As Variant) As String
Const sHEX = "0123456789ABCDEF"
Dim lngRed As Long, lngGreen As Long, lngBlue As Long

  If VarType(Color) = vbLong Then
     lngRed = Color Mod 256&
     lngGreen = (Color Mod 65536) \ 256&
     lngBlue = Color \ 65536
  ElseIf VarType(Color) = vbString Then
     Color = Right$(Color, 6) '// Eksempel: #D0D5DF
     lngRed = 16& * (InStr(1, sHEX, Mid$(Color, 1, 1), vbTextCompare) - 1) + 1& * (InStr(1, sHEX, Mid$(Color, 2, 1), vbTextCompare) - 1)
     lngGreen = 16& * (InStr(1, sHEX, Mid$(Color, 3, 1), vbTextCompare) - 1) + 1& * (InStr(1, sHEX, Mid$(Color, 4, 1), vbTextCompare) - 1)
     lngBlue = 16& * (InStr(1, sHEX, Mid$(Color, 5, 1), vbTextCompare) - 1) + 1& * (InStr(1, sHEX, Mid$(Color, 6, 1), vbTextCompare) - 1)
  Else
     Stop
  End If
  fcnGetRTFColor = "\red" & CStr(lngRed) & "\green" & CStr(lngGreen) & "\blue" & CStr(lngBlue) & ";"
End Function

Function HighlightAll()
  rtf1.SelStart = 0
  rtf1.SelLength = Len(rtf1.Text)
  ColorHTML
  rtf1.SelStart = 1
End Function

Function Highlight_Selected()
  rtf1.Visible = False
  ColorHTML
  rtf1.SelLength = 0
  rtf1.Visible = True
End Function

Function Un_Highlight()
  rtf1.TextRTF = rtf1.Text
End Function

Function ColorHTML()
Dim SS As Long
Dim SL As Long
Dim strBSL As String
Dim strESL As String
Dim header As String
Dim colortbl As String
Dim footer As String
Dim rtfcolor(4) As String
Dim tmpstr As String
Dim TagregEx, Match, Matches

  If rtf1.SelLength < 1 Then Exit Function

  rtfcolor(0) = fcnGetRTFColor(varColorText)
  rtfcolor(1) = fcnGetRTFColor(varColorTag)
  rtfcolor(2) = fcnGetRTFColor(varColorProp)
  rtfcolor(3) = fcnGetRTFColor(varColorPropVal)
  rtfcolor(4) = fcnGetRTFColor(varColorComment)
  colortbl = Join(rtfcolor, ";") & ";"
  header = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss MS Sans Serif;}}"
  colortbl = "{\colortbl" & colortbl & "}"
  header = header & vbCrLf & colortbl & vbCrLf & "\deflang1033\pard\plain\f0\fs17 "
  footer = "\par \plain\f2\fs17\cf0" & vbCrLf & "\par }"
  strBSL = ""
  strESL = ""

  If rtf1.SelLength > 0 Then
     SS = rtf1.SelStart
     SL = rtf1.SelLength
     rtf1.SelStart = 0
     rtf1.SelLength = SS
     strBSL = rtf1.SelRTF
     rtf1.SelStart = SS + SL
     rtf1.SelLength = Len(rtf1.Text) - rtf1.SelStart
     strESL = rtf1.SelRTF
     rtf1.SelStart = SS
     rtf1.SelLength = SL
  End If

  tmpstr = rtf1.SelText
  tmpstr = ReplaceText("([{}\\])", "\$1", tmpstr)
  tmpstr = ReplaceText("(\r)", "\par \r", tmpstr)
  tmpstr = ReplaceText("(<[^>]+>)", "\plain\f2\fs17\cf1 $1\plain\f2\fs17\cf0 ", tmpstr)
  tmpstr = ReplaceText("( \w[\w\d\s:_\-\.]* *= *)(""[^""]+""|'[^']+'|\d+)", "\plain\f2\fs17\cf2 $1\plain\f2\fs17\cf3 $2\plain\f2\fs17\cf1 ", tmpstr)
  rtf1.TextRTF = header & strBSL & tmpstr & "\plain\f2\fs17\cf0 " & strESL & footer
  rtf1.SelStart = SS
  rtf1.SelLength = SL

  Set TagregEx = New RegExp

  TagregEx.Pattern = ">[^<]*=[^>]*<"
  TagregEx.IgnoreCase = False
  TagregEx.Global = True

  Set Matches = TagregEx.Execute(rtf1.SelText)

  For Each Match In Matches
     rtf1.SelStart = Match.FirstIndex + SS + 1
     rtf1.SelLength = Match.Length - 2
     rtf1.SelColor = vbBlack
  Next

  Set TagregEx = New RegExp

  TagregEx.Pattern = "<!--[\w\W]+?-->"
  TagregEx.IgnoreCase = False
  TagregEx.Global = True

  Set Matches = TagregEx.Execute(rtf1.SelText)

  For Each Match In Matches
     rtf1.SelStart = Match.FirstIndex + SS
     rtf1.SelLength = Match.Length
     rtf1.SelColor = &H808080
  Next
End Function

Function ReplaceText(patrn, replStr, textStr)
Dim regEx, str1
Set regEx = New RegExp

  regEx.Pattern = patrn
  regEx.IgnoreCase = True
  regEx.Global = True
  ReplaceText = regEx.Replace(textStr, replStr)
End Function

Private Function INtag() As Boolean
  If rtf1.SelStart > 0 Then
     If InStrRev(rtf1.Text, "<", rtf1.SelStart, vbTextCompare) > InStrRev(rtf1.Text, ">", rtf1.SelStart, vbTextCompare) Then INtag = True
  End If
End Function

Private Function INcomment() As Boolean
  If rtf1.SelStart > 0 Then
     If InStrRev(rtf1.Text, "<!--", rtf1.SelStart, vbTextCompare) > InStrRev(rtf1.Text, "-->", rtf1.SelStart, vbTextCompare) Then INcomment = True
  End If
End Function

Private Function INpropval() As Boolean
Dim x, y As Long
  
  x = InStrRev(rtf1.Text, """", rtf1.SelStart, vbTextCompare)
  y = InStrRev(rtf1.Text, "=", rtf1.SelStart, vbTextCompare)

  If x > y Then
     If InStrRev(rtf1.Text, """", x - 1, vbTextCompare) < InStrRev(rtf1.Text, "=", x - 1, vbTextCompare) Then INpropval = True
  End If
End Function

Private Sub rtf1_Change()
  RaiseEvent Change
End Sub

Private Sub rtf1_KeyPress(KeyAscii As Integer)
  If INcomment = True Then
     Exit Sub
  Else
     If Chr(KeyAscii) = "<" Then rtf1.SelColor = varColorTag
  End If

  If INtag = True Then
        If Chr(KeyAscii) = "-" Then
            If Not Len(rtf1.Text) < 3 Then
                rtf1.SelStart = rtf1.SelStart - 3
                rtf1.SelLength = 3
                Debug.Print rtf1.SelText

                    If rtf1.SelText = "<!-" Then
                        rtf1.SelColor = varColorComment
                    End If
                        rtf1.SelStart = rtf1.SelStart + 4
                    End If
            End If

            If Chr(KeyAscii) = " " Then
                If INpropval Then
                    rtf1.SelColor = varColorPropVal
                Else
                    rtf1.SelColor = varColorProp
                End If

            ElseIf Chr(KeyAscii) = "=" Then
                rtf1.SelText = "="
                rtf1.SelColor = varColorPropVal
                KeyAscii = 0
            ElseIf Chr(KeyAscii) = ">" Then
                rtf1.SelColor = varColorTag
                rtf1.SelText = ">"
                KeyAscii = 0
                rtf1.SelColor = varColorText
            End If

  End If
End Sub

Private Sub rtf1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode & Shift = "1901" Then '// user pressed ">"
     rtf1.SelColor = vbBlack
  End If
End Sub

Function OpenFile(strFilNavn As String)
  rtf1.LoadFile strFilNavn, rtfText
End Function

Function SaveFile(strFilNavn As String)
  rtf1.SaveFile strFilNavn, rtfText
End Function

Private Sub rtf1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     RaiseEvent RightClick
  ElseIf Button = 1 Then
     RaiseEvent LeftClick
  End If
End Sub

Private Sub rtf1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo hell
  RaiseEvent DropFile(Data.Files(1))
hell: Exit Sub
End Sub

Private Sub UserControl_Initialize()
  apppath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")

  varColorText = &H0&
  varColorTag = &HFF0000
  varColorProp = &H800000
  varColorPropVal = &H800080
  varColorComment = &HC0C0C0
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
  rtf1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
