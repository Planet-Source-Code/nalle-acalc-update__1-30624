VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   0  'None
   ClientHeight    =   3660
   ClientLeft      =   3060
   ClientTop       =   705
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   Begin VB.CommandButton cmdBack 
      Caption         =   "OK"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame fraRound 
      Caption         =   "Round setting"
      Height          =   2295
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   2295
      Begin VB.ListBox lstRound 
         Height          =   1425
         ItemData        =   "frmVAT.frx":0000
         Left            =   1440
         List            =   "frmVAT.frx":0019
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblRound 
         BackStyle       =   0  'Transparent
         Caption         =   "How many decimals do You need ? "
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraVAT 
      Caption         =   "Value Added"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2295
      Begin VB.CommandButton cmdOK 
         Caption         =   "Set"
         Height          =   275
         Left            =   1200
         TabIndex        =   8
         Top             =   2280
         Width           =   675
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   275
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2280
         Width           =   675
      End
      Begin VB.TextBox txtVAT 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label lblVAT 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmVAT.frx":0032
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   5520
      Picture         =   "frmVAT.frx":011B
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   5880
      Top             =   480
      Width           =   195
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   0
      Left            =   0
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   1
      Left            =   240
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   2
      Left            =   480
      MousePointer    =   7  'Size N S
      Top             =   0
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   3
      Left            =   480
      MousePointer    =   7  'Size N S
      Top             =   240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   4
      Left            =   1440
      MousePointer    =   8  'Size NW SE
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   5
      Left            =   1680
      MousePointer    =   6  'Size NE SW
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   6
      Left            =   1920
      MousePointer    =   6  'Size NE SW
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   7
      Left            =   2160
      MousePointer    =   8  'Size NW SE
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   5880
      Picture         =   "frmVAT.frx":0365
      Top             =   120
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3960
      TabIndex        =   0
      Top             =   1080
      Width           =   765
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   5520
      Picture         =   "frmVAT.frx":0724
      Top             =   840
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   5520
      Picture         =   "frmVAT.frx":095C
      Top             =   480
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   4080
      Picture         =   "frmVAT.frx":0BA3
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   4440
      Picture         =   "frmVAT.frx":103E
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   4800
      Picture         =   "frmVAT.frx":149C
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   5160
      Picture         =   "frmVAT.frx":1718
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   4080
      Picture         =   "frmVAT.frx":19A5
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   4440
      Picture         =   "frmVAT.frx":1A5E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   4800
      Picture         =   "frmVAT.frx":1AF6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   5160
      Picture         =   "frmVAT.frx":1B84
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim maxreached As Boolean

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    MomsPlus = txtVAT.Text
    MomsMinus = (1 - (100 / (MomsPlus + 100))) * 100
    'save, form, apps name, variable, value(get the value and count it up by 1.
    SaveSetting frmMain.Name, App.Title, "MomsPlus", MomsPlus
    SaveSetting frmMain.Name, App.Title, "MomsMinus", MomsMinus
End Sub

Private Sub Form_Activate()
    txtVAT.SetFocus
    Call Color
End Sub

Private Sub Form_Load()
    MakeWindow Me, True
  '  AlwaysOnTop Me, True
    cmdOK.BackColor = RGB(207, 207, 207)
    cmdClear.BackColor = RGB(207, 207, 207)
    fraVAT.BackColor = RGB(207, 207, 207)
    fraRound.BackColor = RGB(207, 207, 207)
End Sub

Private Sub imgTitleClose_Click()
    Unload Me
End Sub

Private Sub imgTitleHelp_Click()
    MsgBox "Do you realy need help with this ??"
End Sub

Private Sub imgTitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub
Private Sub cmdClear_Click()
' Clears the text box
    txtVAT.Text = ""
    txtVAT.SetFocus
End Sub

Private Sub txtVAT_GotFocus()
    txtVAT.SelLength = Len(txtVAT)
End Sub

Private Sub txtVAT_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 44 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Then
     Else
    KeyAscii = 0
 End If
End Sub

Private Sub txtVAT_KeyUp(KeyCode As Integer, Shift As Integer)
    maxreached = (Len(txtVAT) >= 8)
    If maxreached Then
       cmdOK.SetFocus
        Exit Sub
    End If
End Sub

Private Sub lstRound_Click()
    Rund = Val(lstRound.ListIndex)
    
    SaveSetting frmMain.Name, App.Title, "Rund", Rund


End Sub

Public Sub Color()
Dim ctl As Object
On Error Resume Next
    For Each ctl In Me.Controls
        If TypeOf ctl Is Image Then
            Debug.Print ctl.Name
            ctl.Picture = LoadPicture(App.Path & "\Images\" & ImgCol & "\" & ctl.Name & ".gif")
        End If
    Next
    SetStateBtn Me, Me.WindowState
    For Each ctl In Me.Controls
        ctl.Refresh
    Next

End Sub

