VERSION 5.00
Begin VB.Form frmMenuForm 
   Caption         =   "Menu"
   ClientHeight    =   2625
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRound 
         Caption         =   "&Round setting"
      End
      Begin VB.Menu mnuVAT 
         Caption         =   "&Value added"
      End
      Begin VB.Menu mnuState 
         Caption         =   "&State"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "&Color"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    frmMenuForm.BackColor = RGB(207, 207, 207)
End Sub

Private Sub Form_Unload(p_intCancel As Integer)
    On Error GoTo PROC_ERR
    Set frmMenuForm = Nothing
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox Err.Description
    Resume PROC_EXIT
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
    Unload Me
End Sub

Private Sub mnuColor_Click()
    frmMain.fraColor.Visible = True
    frmMain.lstColor.SetFocus
    Unload Me
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show
    Unload Me
End Sub

Private Sub mnuRound_Click()
    frmSetting.Show
    frmSetting.lstRound.SetFocus
    Unload Me
End Sub

Private Sub mnuState_Click()
    frmMain.fraState.Visible = True
    frmMain.lstState.SetFocus
    Unload Me
End Sub

Private Sub mnuVAT_Click()
    frmSetting.Show
    frmSetting.txtVAT.SetFocus
    Unload Me
End Sub
