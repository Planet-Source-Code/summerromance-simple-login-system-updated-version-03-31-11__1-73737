Attribute VB_Name = "programsettings"
Option Explicit
Public edit As Boolean
Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MINMAXINFO
          ptReserved As POINTAPI
          ptMaxSize As POINTAPI
          ptMaxPosition As POINTAPI
          ptMinTrackSize As POINTAPI
          ptMaxTrackSize As POINTAPI
End Type

Global lpPrevWndProc As Long
Global gHW As Long

Private Declare Function DefWindowProc Lib "user32" Alias _
         "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias _
         "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
          ByVal hwnd As Long, ByVal msg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
         "SetWindowLongA" (ByVal hwnd As Long, _
          ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias _
         "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, _
          ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias _
         "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, _
          ByVal cbCopy As Long)

Public Sub CheckSoftware(x As Form)
On Error GoTo errs
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        MsgBox "Infinity Server is still Running!    ", _
               vbCritical, "Running"
        App.Title = ""
        x.Caption = ""
        AppActivate SaveTitle$
        SendKeys "%{ENTER}", True
        End
    End If
    Exit Sub
errs:
    End
    Exit Sub
End Sub

Public Sub MinAllExceptOne(FormToStay As String)
Dim oFrm As Form
For Each oFrm In Forms
    If oFrm.Name <> FormToStay And Not (TypeOf oFrm Is MDIForm) Then
       oFrm.WindowState = 0
       Set oFrm = Nothing
    Else
       oFrm.WindowState = 2
    End If
Next
End Sub

Public Sub closeAllExceptOne(Optional FormsToStay As String)
Dim oFrms As Form
For Each oFrms In Forms
    If oFrms.Name <> FormsToStay And Not (TypeOf oFrms Is MDIForm) Then
       Unload oFrms
       Set oFrms = Nothing
    End If
Next
End Sub

Public Sub Hook()
          'Start subclassing.
          lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
             AddressOf WindowProc)
End Sub

Public Sub Unhook()
          Dim temp As Long

          'Cease subclassing.
          temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
         ByVal wParam As Long, ByVal lParam As Long) As Long
          Dim MinMax As MINMAXINFO

         'Check for request for min/max window sizes.
          If uMsg = WM_GETMINMAXINFO Then
              'necesary for the caption of an MDI child (when maximized)
              '(Thanks to Marvin Chinchilla for this information)
              WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, _
        wParam, lParam)
              'Retrieve default MinMax settings
              CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

              'Specify new minimum size for window.
              MinMax.ptMinTrackSize.x = (6000 / Screen.TwipsPerPixelX)
              MinMax.ptMinTrackSize.y = (7365 / Screen.TwipsPerPixelY)

              'Specify new maximum size for window.
              MinMax.ptMaxTrackSize.x = Screen.Width
              MinMax.ptMaxTrackSize.y = Screen.Height

              'Copy local structure back.
              CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

              WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
          Else
              WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, _
                 wParam, lParam)
          End If
End Function

