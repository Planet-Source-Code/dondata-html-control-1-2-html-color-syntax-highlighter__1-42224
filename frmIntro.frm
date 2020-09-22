VERSION 5.00
Object = "*\AHTMLControl.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIntro 
   Caption         =   "HTMLControlTest"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7215
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin HTMLControl.txtHTML txtHTML1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2143
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   1200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu Bar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
  txtHTML1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub mnuCopy_Click()
  txtHTML1.Copy
End Sub

Private Sub mnuCut_Click()
  txtHTML1.Cut
End Sub

Private Sub mnuDelete_Click()
  txtHTML1.Delete
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuOpen_Click()
On Error GoTo Hell
  dlgOpen.ShowOpen
  txtHTML1.OpenFile dlgOpen.FileName
  txtHTML1.HighlightAll
Hell: Exit Sub
End Sub

Private Sub mnuPaste_Click()
  txtHTML1.Paste
End Sub

Private Sub mnuSave_Click()
On Error GoTo Hell
  dlgOpen.ShowSave
  txtHTML1.SaveFile dlgOpen.FileName
  txtHTML1.HighlightAll
Hell: Exit Sub
End Sub

Private Sub mnuSelectAll_Click()
  txtHTML1.SelectAll
End Sub

Private Sub mnuUndo_Click()
  MsgBox txtHTML1.CanUndo, vbOKOnly + vbInformation, "Can The Control Undo?"
  txtHTML1.Undo
End Sub

Private Sub txtHTML1_DropFile(FileName As String)
  MsgBox FileName, vbOKOnly + vbInformation, "A File Is Dropped"
End Sub

Private Sub txtHTML1_RightClick()
  PopupMenu mnuEdit
End Sub
