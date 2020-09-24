VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fDelay 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "a"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3315
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton optDummy 
      Caption         =   "Option1"
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   1650
      Width           =   195
   End
   Begin VB.CommandButton btExit 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   270
      Left            =   1380
      TabIndex        =   2
      Top             =   975
      Width           =   570
   End
   Begin MSComctlLib.Slider sliDelay 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   255
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   556
      _Version        =   393216
      LargeChange     =   1
      Max             =   60
      TickFrequency   =   5
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0      5     10    15   20   25   30   35   40    45   50   55   60"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   150
      Index           =   1
      Left            =   150
      TabIndex        =   6
      Top             =   585
      Width           =   3030
   End
   Begin VB.Label lbOff 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Timeout Off"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   1260
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Timeout Delay:"
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   1545
   End
   Begin VB.Label lbMinutes 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   1395
      TabIndex        =   5
      Top             =   750
      Width           =   510
   End
End
Attribute VB_Name = "fDelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btExit_Click()

    Hide

End Sub

Private Sub Form_Load()

    Caption = App.ProductName
    sliDelay_Change

End Sub

Private Sub sliDelay_Change()

    lbOff.Visible = (sliDelay = 0)
    lbMinutes.Visible = sliDelay

End Sub

Private Sub sliDelay_GotFocus()

    optDummy.SetFocus

End Sub

Private Sub sliDelay_Scroll()

    sliDelay_Change

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 15:49)  Decl: 1  Code: 35  Total: 36 Lines
':) CommentOnly: 0 (0%)  Commented: 0 (0%)  Empty: 15 (41,7%)  Max Logic Depth: 1
