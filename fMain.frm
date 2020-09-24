VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   1020
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   930
   ControlBox      =   0   'False
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   930
   StartUpPosition =   2  'Bildschirmmitte
   Visible         =   0   'False
   Begin VB.Menu mnuTray 
      Caption         =   "*"
      Visible         =   0   'False
      Begin VB.Menu mnuTimeout 
         Caption         =   "Set Timeout"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide Menu"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Found at PSC a few days ago and heavily modified - looked for it again just now to give proper credit
'to the original author but can't find it no more. anyway, tnx for showing how it goes.

Public WithEvents Systray   As cSystray
Attribute Systray.VB_VarHelpID = -1
Private Const sTO           As String = "Timeout"

Private Sub AdjTimeout(ByVal Value As Long)

  Dim TOS As String

    TimeOut = Value
    TOS = IIf(TimeOut = 0, "disabled", TimeOut & " minute" & IIf(TimeOut = 1, vbNullString, "s"))
    SaveSetting App.ProductName, sTO, sTO, CStr(TimeOut)
    mnuTimeout.Caption = "Set Timeout (currently " & TOS & ")"
    Systray.Tooltip = App.ProductName & vbCrLf & "Idle Timeout " & TOS & vbCrLf & vbCrLf & "Right click for menu..."
    IdleStopDetection
    If TimeOut Then
        IdleBeginDetection TimeOut
    End If

End Sub

Private Sub Form_Load()

  Dim t     As Date

    Set Systray = New cSystray
    If App.PrevInstance Then
        Beeper 440, 30
        Beeper 220, 10
        Unload Me
      Else 'APP.PREVINSTANCE = FALSE/0
        AdjTimeout GetSetting(App.ProductName, sTO, sTO, 30)
        With Systray
            .SetOwner Me
            .AddIconToTray Icon.Handle, , True
            .ShowBalloon "        Running  in  System Tray", App.ProductName, InfoIcon Or SoundOff
            t = Now
            Do
                DoEvents
            Loop Until Now > t + 0.00008
            .HideBalloon
        End With 'SYSTRAY
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    IdleStopDetection
    Break = True
    With Systray
        If .IsIconInTray Then
            .RemoveIconFromTray
        End If
    End With 'SYSTRAY

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuTimeout_Click()

    With fDelay
        .sliDelay.Value = TimeOut
        .Move Screen.Width - .Width - 420, Screen.Height - .Height - 420
        .Show vbModal
        AdjTimeout .sliDelay.Value
        Unload fDelay
    End With 'FDELAY

End Sub

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 15:49)  Decl: 7  Code: 74  Total: 81 Lines
':) CommentOnly: 2 (2,5%)  Commented: 4 (4,9%)  Empty: 19 (23,5%)  Max Logic Depth: 4
