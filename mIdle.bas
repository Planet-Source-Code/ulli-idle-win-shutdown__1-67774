Attribute VB_Name = "modIdle"
Option Explicit

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, BufferLength As Any, PreviousState As Any, ReturnLength As Any) As Long
Public Declare Function Beeper Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function BeginIdleDetection Lib "Msidle.dll" Alias "#3" (ByVal pfnCallback As Long, ByVal dwIdleMin As Long, ByVal dwReserved As Long) As Long
Private Declare Function EndIdleDetection Lib "Msidle.dll" Alias "#4" (ByVal dwReserved As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As Currency) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Enum ApiConsts
    USER_IDLE_BEGIN = 1
    USER_IDLE_END = 2
    TOKEN_ADJUST_PRIV = &H20
    SE_PRIV_ENABLED = 2
End Enum
#If False Then
Private USER_IDLE_BEGIN, USER_IDLE_END, TOKEN_ADJUST_PRIVILEGES, SE_PRIVILEGE_ENABLED
#End If
Private Const SE_PRIV_NAME  As String = "SeShutdownPrivilege"

Public Enum LogOffMode
    EWX_LOGOFF = 0
    EWX_SHUTDOWN = 1
    EWX_REBOOT = 2
    EWX_FORCE = 4
    EWX_POWEROFF = 8
    EWX_FORCEIFHUNG = 16
End Enum
#If False Then
Private EWX_LOGOFF, EWX_SHUTDOWN, EWX_REBOOT, EWX_FORCE, EWX_POWEROFF, EWX_FORCEIFHUNG
#End If

Public Enum LogOffReturn
    EWX_OK = 0
    EWX_ExitErr = 1
    EWX_NoPrivAdjust = 2
    EWX_PrivNotFound = 3
    EWX_NoTokenHandle = 4
End Enum
#If False Then
Private EWX_OK, EWX_ExitErr, EWX_PrivNotFound, EWX_NoPrivAdjust, EWX_NoTokenHandle
#End If

Private Const EWX_ALL As Long = EWX_LOGOFF Or EWX_SHUTDOWN Or EWX_REBOOT Or EWX_FORCE Or EWX_POWEROFF Or EWX_FORCEIFHUNG

Private Type TOKEN_PRIVILEGES
    PrivCnt     As Long
    pLuid       As Currency
    Attributes  As Long
End Type

Public TimeOut  As Long
Public Break    As Boolean

Public Sub IdleBeginDetection(Optional ByVal IdleMinutes As Long = 60)

    BeginIdleDetection AddressOf IdleCallBack, IdleMinutes, 0&

End Sub

Private Sub IdleCallBack(ByVal dwState As Long)

  Dim i As Long
  Dim j As Long

    Select Case dwState
      Case USER_IDLE_BEGIN
        fMain.Systray.ShowBalloon "        About to shut down...", App.ProductName, WarningIcon Or SoundOff
        Break = False
        i = 10000
        Do Until i < 100 Or Break
            Beeper 1000, 20
            For j = 1 To i / 10
                Sleep 10
                DoEvents
                If Break Then
                    Exit For 'loop varying j
                End If
            Next j
            i = i / 1.5
        Loop
        fMain.Systray.HideBalloon
        If Not Break Then
            '-----------------------------------------
            DoEvents
            ShutDownWindows
            '-----------------------------------------
        End If
      Case USER_IDLE_END
        Break = True
    End Select

End Sub

Public Sub IdleStopDetection()

    EndIdleDetection 0&

End Sub

Public Function ShutDownWindows(Optional ShutDownMode As LogOffMode = EWX_FORCEIFHUNG Or EWX_POWEROFF) As LogOffReturn

  Dim hToken As Long
  Dim TokenPriv As TOKEN_PRIVILEGES

    TokenPriv.PrivCnt = 1 'one privilege
    TokenPriv.Attributes = SE_PRIV_ENABLED 'we want to enable it
    If OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIV, hToken) Then  'get handle for the access token for current process
        If LookupPrivilegeValue(vbNullString, SE_PRIV_NAME, TokenPriv.pLuid) Then  'Get LUID for SHUTDOWN-privilege
            If AdjustTokenPrivileges(hToken, 0&, TokenPriv, ByVal 0&, ByVal 0&, ByVal 0&) Then 'adjust privilege
                If ExitWindowsEx(ShutDownMode And EWX_ALL, 0&) Then  'exit okay
                    ShutDownWindows = EWX_OK
                  Else 'could not exit 'NOT EXITWINDOWSEX(SHUTDOWNMODE...
                    ShutDownWindows = EWX_ExitErr
                End If
              Else 'could not adjust privilege 'NOT ADJUSTTOKENPRIVILEGES(HTOKEN,...
                ShutDownWindows = EWX_NoPrivAdjust
            End If
          Else 'could not fild LUID for privilege 'NOT LOOKUPPRIVILEGEVALUE(VBNULLSTRING,...
            ShutDownWindows = EWX_PrivNotFound
        End If
      Else 'could not open process 'NOT OPENPROCESSTOKEN(GETCURRENTPROCESS,...
        ShutDownWindows = EWX_NoTokenHandle
    End If

End Function

':) Ulli's VB Code Formatter V2.22.14 (2007-Feb-02 15:49)  Decl: 56  Code: 76  Total: 132 Lines
':) CommentOnly: 2 (1,5%)  Commented: 10 (7,6%)  Empty: 21 (15,9%)  Max Logic Depth: 5
