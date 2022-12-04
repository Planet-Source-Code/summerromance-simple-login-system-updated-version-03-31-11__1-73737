Attribute VB_Name = "minTray"
Public Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Public m_IconData As NOTIFYICONDATA
Public Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Public Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
 

Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2
 

Public Const NIIF_NONE = &H0
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_INFO = &H1
 Public Const NIIF_GUID = &H4


Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
 
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Function Popup(Message As String, Title As String)
   With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = mdimain.tryicon.hWnd
        .uID = 1&
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = mdimain.tryicon.Picture
        .szTip = Title & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        .dwInfoFlags = NIIF_WARNING
         End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Function

Public Sub seticon()
With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = mdimain.tryicon.hWnd
        .uID = 1&
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = mdimain.tryicon.Picture
        .szTip = "Sample" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
   
End Sub

Public Sub delicon()
Shell_NotifyIcon NIM_DELETE, m_IconData

End Sub



