VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3810
   ClientLeft      =   6255
   ClientTop       =   495
   ClientWidth     =   7470
   ForeColor       =   &H80000008&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   254
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraState 
      Caption         =   "State setting"
      Height          =   2415
      Left            =   3840
      TabIndex        =   43
      Top             =   1200
      Width           =   2295
      Visible         =   0   'False
      Begin VB.ListBox lstState 
         Height          =   1425
         ItemData        =   "Form1.frx":030A
         Left            =   240
         List            =   "Form1.frx":031A
         TabIndex        =   44
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame fraColor 
      Caption         =   "Color setting"
      Height          =   2415
      Left            =   3240
      TabIndex        =   41
      Top             =   1200
      Width           =   2295
      Visible         =   0   'False
      Begin VB.ListBox lstColor 
         Height          =   1620
         ItemData        =   "Form1.frx":0346
         Left            =   240
         List            =   "Form1.frx":035C
         TabIndex        =   42
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   5640
      Tag             =   "2"
      Top             =   3000
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   210
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3000
      Width           =   480
   End
   Begin VB.CommandButton CopyButton 
      BackColor       =   &H00808080&
      Caption         =   "<"
      Height          =   315
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Copy to clipboard"
      Top             =   360
      Width           =   315
      Visible         =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   480
      Width           =   2295
      Begin VB.Label Readout 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   60
         TabIndex        =   34
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2280
      Width           =   300
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2280
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2280
      Width           =   660
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1920
      Width           =   300
   End
   Begin VB.CommandButton Operator 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1920
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1920
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1920
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1920
      Width           =   300
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1560
      Width           =   300
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1560
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1560
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1560
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1560
      Width           =   300
   End
   Begin VB.CommandButton CancelEntry 
      BackColor       =   &H00808080&
      Caption         =   "CE"
      Height          =   300
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "F3"
      Top             =   1200
      Width           =   300
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00808080&
      Caption         =   "C"
      Height          =   300
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "F2 or Del"
      Top             =   1200
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1200
      Width           =   300
   End
   Begin VB.CommandButton NumKey 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1200
      Width           =   300
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "M+"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   3720
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "M-"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   3720
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "MR"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   4200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.CommandButton MemoKey 
      Caption         =   "MC"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   4200
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   420
      Visible         =   0   'False
   End
   Begin VB.ListBox lstKvitto 
      Height          =   2205
      Left            =   2685
      TabIndex        =   10
      Top             =   450
      Width           =   2100
   End
   Begin VB.CommandButton cmdEjKvitto 
      Caption         =   "<<   No monitor"
      Height          =   210
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "F6"
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   210
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "P"
      Top             =   3375
      Width           =   1275
   End
   Begin VB.CommandButton cmdKvitto 
      Caption         =   "Monitor   >>"
      Height          =   210
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "F5"
      Top             =   3000
      Width           =   1275
   End
   Begin VB.CommandButton cmdProc 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "F4"
      Top             =   2280
      Width           =   300
   End
   Begin VB.CommandButton cmdSlagremsa 
      Caption         =   "Readout receipt"
      Height          =   210
      Left            =   3015
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "F7"
      Top             =   3135
      Width           =   1275
   End
   Begin VB.CommandButton cmdTillbaka 
      Caption         =   "Back to monitor"
      Height          =   210
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint2 
      Caption         =   "&Print"
      Height          =   210
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "P"
      Top             =   3480
      Width           =   1275
   End
   Begin VB.CommandButton cmdMuPlus 
      Caption         =   "VA +"
      Height          =   300
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "F11 or Page Up"
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton cmdMuMinus 
      Caption         =   "VA -"
      Height          =   300
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "F12 or Page Down"
      Top             =   2640
      Width           =   915
   End
   Begin VB.Label lblRound 
      BackStyle       =   0  'Transparent
      Caption         =   "Round ="
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   46
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblDec 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   45
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label lblAlarm 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   40
      Top             =   3360
      Width           =   630
   End
   Begin VB.Label lblAlarmset 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Alarm set"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   990
      TabIndex        =   39
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   38
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lblMemoFlag 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      ToolTipText     =   "If M, memory  not zero"
      Top             =   120
      Width           =   225
      Visible         =   0   'False
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      X1              =   169
      X2              =   169
      Y1              =   32
      Y2              =   250
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   5280
      Picture         =   "Form1.frx":0388
      Top             =   360
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   5520
      Picture         =   "Form1.frx":0747
      Top             =   360
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   7320
      Top             =   720
      Width           =   195
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   7
      Left            =   2400
      MousePointer    =   8  'Size NW SE
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   6
      Left            =   2160
      MousePointer    =   6  'Size NE SW
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   5
      Left            =   1920
      MousePointer    =   6  'Size NE SW
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   4
      Left            =   1680
      MousePointer    =   8  'Size NW SE
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   3
      Left            =   720
      MousePointer    =   7  'Size N S
      Top             =   360
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   2
      Left            =   720
      MousePointer    =   7  'Size N S
      Top             =   120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   1
      Left            =   480
      MousePointer    =   9  'Size W E
      Top             =   120
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   0
      Left            =   240
      MousePointer    =   9  'Size W E
      Top             =   120
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator"
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
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0991
      Top             =   480
      Width           =   195
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0BDB
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   7320
      Picture         =   "Form1.frx":0E25
      Top             =   0
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   5880
      Picture         =   "Form1.frx":106F
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   6240
      Picture         =   "Form1.frx":17B9
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "Form1.frx":1F03
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   6960
      Picture         =   "Form1.frx":264D
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5880
      Picture         =   "Form1.frx":2D97
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   6240
      Picture         =   "Form1.frx":34E1
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "Form1.frx":3C2B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   6960
      Picture         =   "Form1.frx":4375
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Temp
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI
'Original math functions authored by Herman Lui
'I added some functions:
'Listbox and print
'Helpfunction
'Functionskeys
'Percentage-key
'VAT-key in Sweden VAT is 25 %
'Possibility to change layout during runtime, se Status.
'Always on top
'Some smaller changes
'Have a nice day
'Alarmhandling
Public Alarmtime As String
'Public Alarmtime2 As String
'Right adjustment listbox
Private Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Must be a long or integer array
Dim mTabs(0) As Long
Dim pTabs(0) As Long
'Originaldeklaration
Const Maxdigits = 13        ' After this, scientific notation
Dim Op1 As Variant          ' Prev input operand
Dim Op2 As Variant          ' Further prev input operand
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Numkey of operands, 0 to 2
Dim LastInput As String     ' Indicate type of last keypress event.
Dim OpFlag As String        ' Indicate pending operation.
Dim PrevReadout As String   ' For restore if "CE"
Dim MemoResult              ' Store result for memo keys
Dim XReadout As String
Dim XOp1 As Variant
Dim XOp2 As Variant
Dim XDecimalFlag As Integer
Dim XNumOps As Integer
Dim XLastInput As String
Dim XOpFlag As String
Dim XCaption As String
Dim XMemoResult
Dim CopyReadout As String
Dim KvittoFlag As String
Dim strTempreadout As String
Dim MinStatus As String
Dim Index As Integer
Dim KnappStatus As Integer
Dim PrevLastInput As String

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText CopyReadout
End Sub



Private Sub Form_Activate()
    Call Counter(Me)
    Dim X As String
'X$ = GetSetting(Form, App, Variable)
    X$ = GetSetting(Me.Name, App.Title, "TimesOpen")
'Set variables
    Rund = GetSetting(Me.Name, App.Title, "Rund")
    MomsPlus = GetSetting(Me.Name, App.Title, "MomsPlus")
    MomsMinus = GetSetting(Me.Name, App.Title, "MomsMinus")
' Set startupinterface
    MinStatus = GetSetting(Me.Name, App.Title, "MinStatus")
    Call Status(MinStatus)
    ResetStatus
'set color
    ImgCol = GetSetting(Me.Name, App.Title, "ImgCol")
    Call Color
    cmdMuPlus.Caption = "VA+ " & Round(MomsPlus, Rund)
    cmdMuMinus.Caption = "VA- " & Round(MomsMinus, Rund)
    lblDec.Caption = Rund
End Sub

Private Sub Form_Load()
'Make "the Form"
    MakeWindow Me, True
    AlwaysOnTop Me, True
' Make the Maximize/Restore button have the Maximize image
   imgTitleMaxRestore.Picture = imgTitleMaximize.Picture
'Alignment wright
'Set the a tab stop to a negative number
    mTabs(0) = -65
    pTabs(0) = -100
'Start clock and set alarm
    lblTime.Caption = Format$(Now, "hh:mm:ss")
    lblAlarm.Caption = "Off"
    fraState.BackColor = RGB(207, 207, 207)
    fraColor.BackColor = RGB(207, 207, 207)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting frmMain.Name, App.Title, "Minstatus", MinStatus

End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    AlwaysOnTop Me, False
    
    PopupMenu frmMenuForm.mnuFile, , (Frame1.Left), (Frame1.Top + Frame1.Height)

End Sub


Private Sub imgTitleClose_Click()
    Unload Me
    End
End Sub

Private Sub imgTitleHelp_Click()
    frmHelp.Show
End Sub

Private Sub imgTitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMaxRestore_Click()
    ChangeState Me
End Sub

Private Sub imgTitleMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub imgTitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub lblAlarm_Change()
    Alarmtime = lblAlarm.Caption
End Sub

Private Sub lblAlarm_Click()
    frmSetAlarm.Show
End Sub

Private Sub lblAlarmset_Click()
    frmSetAlarm.Show
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub lstColor_Click()
ImgCol = lstColor.Text
Call Color
SaveSetting frmMain.Name, App.Title, "ImgCol", ImgCol
fraColor.Visible = False
End Sub


Private Sub lstState_Click()
If lstState.List(lstState.ListIndex) = "Mini" Then
    Call Status("Mini")
Else
If lstState.List(lstState.ListIndex) = "Monitor" Then
    Call Status("Monitor")
Else
If lstState.List(lstState.ListIndex) = "ReadoutReceipt" Then
    Call Status("ReadoutReceipt")
Else
If lstState.List(lstState.ListIndex) = "Receipt" Then
    Call Status("Receipt")
End If
End If
End If
End If

    fraState.Visible = False
End Sub

Private Sub Resizer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Temp = GetCursorPos(OldCursorPos)
End Sub

Private Sub Resizer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Temp = GetCursorPos(NewCursorPos)
    ResizeForm Me, OldCursorPos, NewCursorPos, Index
End Sub

Private Sub imgTitleMain_DblClick()
    ChangeState Me
End Sub

Private Sub cmdEjKvitto_Click()
    Call Status("Mini")
End Sub

Private Sub cmdKvitto_Click()
    Call Status("Monitor")
End Sub

Private Sub cmdMuMinus_Click()
On Error GoTo endcmdMuMinus '--- Just in case!
Dim Moms As String
Dim moms1 As String    'Used Round to two decimal places!
KnappStatus = 3
    Operator_Click 4
    Moms = Readout * (MomsMinus / 100)
    'Moms = Readout * 0.2
    moms1 = Round(Moms, Rund) 'Used Round to two decimal places!
    Readout = Readout - moms1  '* 0.8
    Readout = Round(Readout, Rund)
'Aligment wright
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
    lstKvitto.AddItem vbTab & "vat   - " + moms1
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem " "
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True
SendMessage frmPrint.lstKvitto.hwnd, LB_SETTABSTOPS, 1, pTabs(0)
    frmPrint.lstKvitto.AddItem vbTab & "vat   - " + moms1
    frmPrint.lstKvitto.AddItem vbTab & "= " + Readout
    frmPrint.lstKvitto.AddItem " "
Exit Sub
endcmdMuMinus:
Cancel_Click
Readout = 0

End Sub

Private Sub cmdMuPlus_Click()
On Error GoTo endcmdMuPlus
'Used Round to two decimal places!
Dim Moms As String
KnappStatus = 3
    Operator_Click 4
    Moms = Readout * (MomsPlus / 100)
    Readout = Readout * (1 + (MomsPlus / 100))
    Readout = Round(Readout, Rund)
    Moms = Round(Moms, Rund)
'Aligment wright
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
    lstKvitto.AddItem vbTab & "vat   + " + Moms
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem " "
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True
SendMessage frmPrint.lstKvitto.hwnd, LB_SETTABSTOPS, 1, pTabs(0)
    frmPrint.lstKvitto.AddItem vbTab & "vat   + " + Moms
    frmPrint.lstKvitto.AddItem vbTab & "= " + Readout
    frmPrint.lstKvitto.AddItem " "
Exit Sub
endcmdMuPlus:
Cancel_Click
Readout = 0

End Sub

Private Sub cmdPrint_Click()
       frmPrint.lstKvitto.Appearance = 0
       frmPrint.PrintForm
End Sub

Private Sub cmdPrint2_Click()
       frmPrint.lstKvitto.Appearance = 0
       frmPrint.PrintForm
End Sub

Private Sub cmdProc_Click()
KnappStatus = 3
    Operator_Click 4
            If XOpFlag = "*" Then
                Readout = Round((Readout / 100), Rund)
                Else
            If XOpFlag = "-" Then
                Readout = Round((XOp1 - ((XOp1 * 10) / 100)), Rund)
                Else
            If XOpFlag = "+" Then
                Readout = Round((XOp1 + ((XOp1 * 10) / 100)), Rund)
                Else
            If XOpFlag = "/" Then
                Readout = Round((Readout * 100), Rund)
                Else
            End If
            End If
            End If
            End If
    
'Aligment wright
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
    lstKvitto.AddItem vbTab & "% "
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem " "
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True
SendMessage frmPrint.lstKvitto.hwnd, LB_SETTABSTOPS, 1, pTabs(0)
    frmPrint.lstKvitto.AddItem vbTab & "% "
    frmPrint.lstKvitto.AddItem vbTab & "= " + Readout
    frmPrint.lstKvitto.AddItem " "
End Sub

Private Sub cmdSlagremsa_Click()
    Call Status("ReadoutReceipt")
End Sub

Private Sub cmdTillbaka_Click()
    Call Status("Monitor")
End Sub


Sub ResetStatus()
    Readout = Format(0, "0.")
    PrevReadout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    lblMemoFlag.Caption = " "
    MemoResult = 0
End Sub

Sub RestoreStatus()
    Readout = XReadout
    Op1 = XOp1
    Op2 = XOp2
    DecimalFlag = XDecimalFlag
    NumOps = XNumOps
    LastInput = XLastInput
    OpFlag = XOpFlag
    lblMemoFlag.Caption = XCaption
    MemoResult = XMemoResult
End Sub


Sub MarkStatus()
    XReadout = Readout
    XOp1 = Op1
    XOp2 = Op2
    XDecimalFlag = DecimalFlag
    XNumOps = NumOps
    XLastInput = LastInput
    XOpFlag = OpFlag
    XCaption = lblMemoFlag.Caption
    XMemoResult = MemoResult
End Sub

Private Function maxreached() As Boolean
    maxreached = (Len(Readout) >= Maxdigits)
End Function

Function HasDecimal(strToRead As String) As Boolean
    HasDecimal = InStr(1, strToRead, ".")
End Function

' Copy the "Label" Caption onto the Clipboard.
Private Sub CopyButton_Click()
    Clipboard.SetText Readout
End Sub


Private Sub Cancel_Click()
    ResetStatus
    lstKvitto.Clear
    frmPrint.lstKvitto.Clear
    Operator(4).SetFocus
End Sub


Private Sub CancelEntry_Click()
    RestoreStatus
    LastInput = "CE"
    Operator(4).SetFocus
End Sub

Private Sub cmdDecimal_Click()
    If HasDecimal(Readout) Then             ' One is enough
        Exit Sub
    End If
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        If Len(Readout) = Maxdigits Then
            AlwaysOnTop Me, False
            MsgBox "Maximal digits " & Str(Maxdigits - 1) + _
                vbCrLf & "Try again", , "  Caculator"
            AlwaysOnTop Me, True
            Operator(4).SetFocus
            Exit Sub
        End If
    End If
    
    MarkStatus
    
    If LastInput = "NEG" Then
        If Abs(Val(Readout)) <> 0 Then
            Readout = Format(0, "-0.")
        End If
    ElseIf LastInput <> "NUMS" And LastInput <> "DIGI" Then
        Readout = Format(0, "0.")
    End If
    
    DecimalFlag = True
    LastInput = "DIGI"
    
    If maxreached Then
        AlwaysOnTop Me, False
        MsgBox "Max digits " & Str(Maxdigits - 1) + _
           vbCrLf & " Try again", , "  Calculator"
    AlwaysOnTop Me, True
        RestoreStatus
        Exit Sub
    End If
    Operator(4).SetFocus
End Sub
Private Sub Numkey_Click(Index As Integer)
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        If maxreached Then
            AlwaysOnTop Me, False
            MsgBox "Max digits " & Str(Maxdigits - 1) + _
               vbCrLf & "Try again", , "  Calculator"
                   AlwaysOnTop Me, True
            Operator(4).SetFocus
            Exit Sub
        End If
    End If
    
    MarkStatus
    If LastInput <> "NUMS" And LastInput <> "DIGI" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + NumKey(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + NumKey(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then
        Readout = "-" & Readout
    End If
    LastInput = "NUMS"
  KnappStatus = 1
    Operator(4).SetFocus
End Sub

Private Sub Operator_Click(Index As Integer)
    MarkStatus
    
    strTempreadout = Readout
    
    If LastInput = "NONE" Then
        AlwaysOnTop Me, False
    MsgBox "Start with a number", 48, "  Calculator"
        AlwaysOnTop Me, True
    Exit Sub
    End If
    
    If LastInput = "NUMS" Or LastInput = "DIGI" Then
        NumOps = NumOps + 1
    End If

If OpFlag = "=" Then
KvittoFlag = " "
Else
   KvittoFlag = OpFlag + " "
End If
    
    Select Case NumOps
        Case 0
            If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
                If Abs(Val(Readout)) <> 0 Then
                    Readout = "-" & Readout
                    LastInput = "NEG"
                End If
            End If
        Case 1
            Op1 = Readout
            If Operator(Index).Caption = "-" And (LastInput <> "NUMS" _
                    And LastInput <> "DIGI") And OpFlag <> "=" Then
                If Abs(Val(Readout)) <> 0 Then
                    Readout = "-"
                    LastInput = "NEG"
                End If
            End If
        Case 2
            Op2 = strTempreadout
            Select Case OpFlag
                Case "+"
                    Op1 = CDbl(Op1) + CDbl(Op2)
                Case "-"
                    Op1 = CDbl(Op1) - CDbl(Op2)
                Case "*"
                    Op1 = CDbl(Op1) * CDbl(Op2)
                Case "/"
                    If Op2 = 0 Then
                        AlwaysOnTop Me, False
                        MsgBox "Division by zero not possible", 48, "  Calculator"
                        AlwaysOnTop Me, True

                       RestoreStatus
                       Exit Sub
                    Else
                       Op1 = CDbl(Op1) / CDbl(Op2)
                    End If
               Case "="
                    Op1 = CDbl(Op2)
             End Select
             Readout = Round(Op1, Rund)
             NumOps = 1
             
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
    
Call Kvitto
Call PrintReceipt
KnappStatus = 2
Operator(4).SetFocus
End Sub
Private Sub MemoKey_Click(Index As Integer)
    MarkStatus
    Select Case Index
       Case 0                    ' Memory Plus
            MemoResult = MemoResult + Val(Readout)
       Case 1                    ' Memory Minus
            MemoResult = MemoResult - Val(Readout)
       Case 2                    ' Memory Recall
            Dim s As String
            s = Str(MemoResult)
            If Not HasDecimal(Str(s)) Then
                s = s + "."
            End If
            Readout = s
       Case 3                    ' Memory Clear
            MemoResult = 0
    End Select
     ' Our system is, if MemoResult is not cleared, show "M"
    If MemoResult <> 0 Then
         lblMemoFlag.Caption = "M"
    Else
         lblMemoFlag.Caption = " "
    End If
    
    LastInput = "OPS"
    NumOps = 1
    Op1 = Readout
    Op2 = 0
    Operator(4).SetFocus
End Sub
' Detect keyboard key
Private Sub Form_KeyPress(KeyAscii As Integer)
    MarkStatus
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 46 And KeyAscii <> 43 And _
           KeyAscii <> 45 And KeyAscii <> 42 And _
           KeyAscii <> 47 And KeyAscii <> 61 And _
           KeyAscii <> 13 Then
               KeyAscii = 0
        Else
           Select Case KeyAscii
             Case 46                   ' "."
               cmdDecimal_Click
             Case 43
               Operator_Click (0)      ' re Property "+"
             Case 45                   ' "-"
               Operator_Click (1)
             Case 42                   ' "*"
               Operator_Click (2)
             Case 47                   ' "/"
               Operator_Click (3)
             Case 61                   ' "="
               Operator_Click (4)
             Case 13                   ' As "=" (if Windows allows Enter)
               Operator_Click (4)
           End Select
        End If
    Else
        Numkey_Click (Val(Chr(KeyAscii)))
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyDecimal
    cmdDecimal_Click
Case vbKeyDelete
    Cancel_Click
Case vbKeyEscape
    imgTitleClose_Click
Case vbKeyEnd
    imgTitleClose_Click
Case vbKeyF1
    imgTitleHelp_Click
Case vbKeyF2
    Cancel_Click
Case vbKeyF3
    CancelEntry_Click
Case vbKeyF4
    cmdProc_Click
Case vbKeyHome
    cmdProc_Click
Case vbKeyF5
    cmdKvitto_Click
Case vbKeyF6
    cmdEjKvitto_Click
Case vbKeyF7
    cmdSlagremsa_Click
Case vbKeyF8
    Call Status("Receipt")
Case vbKeyF9
    CopyButton_Click
Case vbKeyF10
    frmAbout.Show
Case vbKeyF11
    cmdMuPlus_Click
Case vbKeyPageUp
    cmdMuPlus_Click
Case vbKeyF12
    cmdMuMinus_Click
Case vbKeyPageDown
    cmdMuMinus_Click
Case 16 And 53
    cmdProc_Click

End Select
End Sub

Private Sub Kvitto()
'Alignment right
SendMessage lstKvitto.hwnd, LB_SETTABSTOPS, 1, mTabs(0)
If KnappStatus = 2 And OpFlag = "=" Then
    lstKvitto.AddItem vbTab & "= " + Readout
   Else
If KnappStatus = 2 Then
    KvittoFlag = "= "
    Else
If KnappStatus = 3 Then
    lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
    Else
If PrevLastInput = "NEG" And OpFlag = "=" Then
    lstKvitto.AddItem vbTab & strTempreadout
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem vbTab & "  "
    Else
If LastInput = "OPS" And OpFlag = "=" Then
    lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
    lstKvitto.AddItem vbTab & "= " + Readout
    lstKvitto.AddItem vbTab & "  "
    Else
    lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
End If
End If
End If
End If
End If
CopyReadout = Readout
PrevLastInput = LastInput
    lstKvitto.Selected(lstKvitto.ListCount - 1) = True
End Sub
Private Sub PrintReceipt()
'Alignment wright
SendMessage frmPrint.lstKvitto.hwnd, LB_SETTABSTOPS, 1, pTabs(0)
If KnappStatus = 2 And OpFlag = "=" Then
    frmPrint.lstKvitto.AddItem vbTab & "= " + Readout
   Else
If KnappStatus = 2 Then
    KvittoFlag = "= "
    Else
If KnappStatus = 3 Then
    frmPrint.lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
    Else
If PrevLastInput = "NEG" And OpFlag = "=" Then
    frmPrint.lstKvitto.AddItem vbTab & strTempreadout
    frmPrint.lstKvitto.AddItem vbTab & "= " + Readout
    frmPrint.lstKvitto.AddItem vbTab & "  "
    Else
If LastInput = "OPS" And OpFlag = "=" Then
    frmPrint.lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
    frmPrint.lstKvitto.AddItem vbTab & "= " + Readout
    frmPrint.lstKvitto.AddItem vbTab & "  "
    Else
    frmPrint.lstKvitto.AddItem vbTab & KvittoFlag + strTempreadout
End If
End If
End If
End If
End If
CopyReadout = Readout
PrevLastInput = LastInput
  '  frmPrint.lstKvitto.Selected(lstKvitto.ListCount - 1) = True
End Sub

Private Sub Status(AppStatus As String)
MinStatus = AppStatus
Select Case AppStatus
Case "Monitor"
    Operator(1).Visible = True
    Operator(3).Visible = True
    Operator(4).Left = 128
    cmdCopy.Top = 200
    cmdCopy.Left = 16
    Frame1.Top = 32
    Frame1.Left = 8
    frmMain.Height = 3810
    frmMain.Width = 5000
    lstKvitto.Height = 147
    lstKvitto.Left = 176
    lstKvitto.Top = 40
    cmdDecimal.Visible = True
    Cancel.Visible = True
    CancelEntry.Visible = True
    cmdProc.Visible = True
    cmdKvitto.Visible = False
    cmdEjKvitto.Visible = True
    cmdPrint.Visible = True
    cmdTillbaka.Visible = False
    cmdPrint2.Visible = False
    Line1.Visible = True
    Readout.Visible = True
    Frame1.Visible = True
    lstKvitto.Visible = True
    cmdMuPlus.Visible = True
    cmdMuMinus.Visible = True
    For Index = 0 To 9
    NumKey(Index).Visible = True
    Next
    For Index = 0 To 4
    Operator(Index).Visible = True
    Next
Case "Mini"
    Operator(1).Visible = True
    Operator(3).Visible = True
    Operator(4).Left = 128
    cmdCopy.Top = 200
    cmdCopy.Left = 16
    Frame1.Top = 32
    Frame1.Left = 8
    frmMain.Height = 3810
    frmMain.Width = 2575
    cmdDecimal.Visible = True
    Cancel.Visible = True
    CancelEntry.Visible = True
    cmdProc.Visible = True
    cmdKvitto.Visible = True
    cmdKvitto.Left = 64
    cmdEjKvitto.Visible = True
    cmdPrint.Visible = True
    cmdTillbaka.Visible = False
    cmdPrint2.Visible = False
    Line1.Visible = False
    Readout.Visible = True
    Frame1.Visible = True
    lstKvitto.Visible = False
    cmdMuPlus.Visible = True
    cmdMuMinus.Visible = True
    For Index = 0 To 9
    NumKey(Index).Visible = True
    Next
    For Index = 0 To 4
    Operator(Index).Visible = True
    Next
    
Case "ReadoutReceipt"
    Operator(1).Visible = False
    Operator(3).Visible = False
    Operator(4).Left = 1000
    cmdCopy.Top = 455
    cmdCopy.Left = 40
    cmdCopy.Width = 40
    lstKvitto.Height = 360
    lstKvitto.Left = 10
    lstKvitto.Top = 75
    lstKvitto.Visible = True
    frmMain.Height = 7200
    frmMain.Width = 2500
    cmdDecimal.Visible = False
    Cancel.Visible = False
    CancelEntry.Visible = False
    cmdProc.Visible = False
    cmdKvitto.Visible = False
    cmdEjKvitto.Visible = False
    cmdPrint.Visible = False
    cmdTillbaka.Visible = True
    cmdTillbaka.Top = 438
    cmdTillbaka.Left = 40
    cmdPrint2.Visible = True
    cmdPrint2.Top = 455
    cmdPrint2.Left = 85
    cmdPrint2.Width = 40
    Line1.Visible = False
    Frame1.Visible = True
    cmdMuPlus.Visible = False
    cmdMuMinus.Visible = False
    lstKvitto.SetFocus
    For Index = 0 To 9
    NumKey(Index).Visible = False
    Next
    For Index = 0 To 3
    Operator(Index).Visible = False
    Next
Case "Receipt"
    Operator(1).Visible = False
    Operator(3).Visible = False
    Operator(4).Left = 1000
    cmdCopy.Top = 455
    cmdCopy.Left = 40
    cmdCopy.Width = 40
    lstKvitto.Height = 400
    lstKvitto.Left = 10
    lstKvitto.Top = 35
    frmMain.Height = 7200
    frmMain.Width = 2395
    cmdDecimal.Visible = False
    Cancel.Visible = False
    CancelEntry.Visible = False
    cmdProc.Visible = False
    cmdKvitto.Visible = False
    cmdEjKvitto.Visible = False
    cmdPrint.Visible = False
    cmdTillbaka.Visible = True
    cmdTillbaka.Top = 438
    cmdTillbaka.Left = 40
    cmdPrint2.Visible = True
    cmdPrint2.Top = 455
    cmdPrint2.Left = 85
    cmdPrint2.Width = 40
    Line1.Visible = False
    Frame1.Visible = False
    lstKvitto.Visible = True
    cmdMuPlus.Visible = False
    cmdMuMinus.Visible = False
    lstKvitto.SetFocus
    For Index = 0 To 9
    NumKey(Index).Visible = False
    Next
    For Index = 0 To 3
    Operator(Index).Visible = False
    Next
End Select
MakeWindow Me, True

End Sub
Private Function Counter(TheForm As Form)
    'save, form, apps name, variable, value(get the value and count it up by 1.
    SaveSetting TheForm.Name, App.Title, "TimesOpen", Val(GetSetting(TheForm.Name, App.Title, "TimesOpen")) + 1
End Function

Private Sub tmrTime_Timer()
    lblTime.Caption = Format$(Now, "hh:mm:ss")
' This happens every 1000 milsec or 1 sec.
    lblTime.Caption = Time ' Setting time
    If Time >= Alarmtime Then
        frmAlarm.Show
    End If
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
Select Case ImgCol
Case "Blue"
    Frame1.ForeColor = &HFF0000
    lblAlarmset.ForeColor = &HFF0000
Case "Green"
    Frame1.ForeColor = &H8000&
    lblAlarmset.ForeColor = &H8000&
Case "Purple"
    Frame1.ForeColor = &HC000C0
    lblAlarmset.ForeColor = &HC000C0
Case "Red"
    Frame1.ForeColor = &HFF&
    lblAlarmset.ForeColor = &HFF&
Case "Grey"
    Frame1.ForeColor = &HFF0000
    lblAlarmset.ForeColor = &HFF0000
Case "Canvas"
    Frame1.ForeColor = &HFF0000
    lblAlarmset.ForeColor = &HFF0000
End Select
    
End Sub
