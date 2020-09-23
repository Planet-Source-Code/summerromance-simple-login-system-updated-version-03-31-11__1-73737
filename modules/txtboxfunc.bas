Attribute VB_Name = "txtboxfunc"
Option Explicit

Public Enum GetWindowLongIndexes
    GWL_EXSTYLE = (-20)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_USERDATA = (-21)
    GWL_WNDPROC = (-4)
End Enum

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As GetWindowLongIndexes) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As GetWindowLongIndexes, ByVal dwNewLong As Long) As Long

Public Enum WindowMessages
    WM_NULL = &H0   ' 0
    WM_CREATE = &H1 ' 1
    WM_DESTROY = &H2    ' 2
    WM_MOVE = &H3   ' 3
    WM_SIZE = &H5   ' 5
    WM_ACTIVATE = &H6   ' 6
    WM_SETFOCUS = &H7   ' 7
    WM_KILLFOCUS = &H8  ' 8
    WM_ENABLE = &HA ' 10
    WM_SETREDRAW = &HB  ' 11
    WM_SETTEXT = &HC    ' 12
    WM_GETTEXT = &HD    ' 13
    WM_GETTEXTLENGTH = &HE  ' 14
    WM_PAINT = &HF  ' 15
    WM_CLOSE = &H10 ' 16
    WM_QUERYENDSESSION = &H11   ' 17
    WM_QUIT = &H12  ' 18
    WM_QUERYOPEN = &H13 ' 19
    WM_ERASEBKGND = &H14    ' 20
    WM_SYSCOLORCHANGE = &H15    ' 21
    WM_ENDSESSION = &H16    ' 22
    WM_SHOWWINDOW = &H18    ' 24
    WM_SETTINGCHANGE = &H1A ' 26
    WM_WININICHANGE = &H1A  ' 26
    WM_DEVMODECHANGE = &H1B ' 27
    WM_ACTIVATEAPP = &H1C   ' 28
    WM_FONTCHANGE = &H1D    ' 29
    WM_TIMECHANGE = &H1E    ' 30
    WM_CANCELMODE = &H1F    ' 31
    WM_CAPTURECHANGED = &H1F    ' 31
    WM_SETCURSOR = &H20 ' 32
    WM_MOUSEACTIVATE = &H21 ' 33
    WM_CHILDACTIVATE = &H22 ' 34
    WM_QUEUESYNC = &H23 ' 35
    WM_GETMINMAXINFO = &H24 ' 36
    WM_PAINTICON = &H26 ' 38
    WM_ICONERASEBKGND = &H27    ' 39
    WM_NEXTDLGCTL = &H28    ' 40
    WM_SPOOLERSTATUS = &H2A ' 42
    WM_DRAWITEM = &H2B  ' 43
    WM_MEASUREITEM = &H2C   ' 44
    WM_DELETEITEM = &H2D    ' 45
    WM_VKEYTOITEM = &H2E    ' 46
    WM_CHARTOITEM = &H2F    ' 47
    WM_SETFONT = &H30   ' 48
    WM_GETFONT = &H31   ' 49
    WM_SETHOTKEY = &H32 ' 50
    WM_GETHOTKEY = &H33 ' 51
    WM_QUERYDRAGICON = &H37 ' 55
    WM_COMPAREITEM = &H39   ' 57
    WM_COMPACTING = &H41    ' 65
    WM_WINDOWPOSCHANGING = &H46 ' 70
    WM_WINDOWPOSCHANGED = &H47  ' 71
    WM_POWER = &H48 ' 72
    WM_COPYDATA = &H4A  ' 74
    WM_CANCELJOURNAL = &H4B ' 75
    WM_INPUTLANGCHANGEREQUEST = &H50    ' 80
    WM_INPUTLANGCHANGE = &H51   ' 81
    WM_HELP = &H53  ' 83
    WM_CONTEXTMENU = &H7B   ' 123
    WM_NCCREATE = &H81  ' 129
    WM_NCDESTROY = &H82 ' 130
    WM_NCCALCSIZE = &H83    ' 131
    WM_NCHITTEST = &H84 ' 132
    WM_NCPAINT = &H85   ' 133
    WM_NCACTIVATE = &H86    ' 134
    WM_GETDLGCODE = &H87    ' 135
    WM_NCMOUSEMOVE = &HA0   ' 160
    WM_NCLBUTTONDOWN = &HA1 ' 161
    WM_NCLBUTTONUP = &HA2   ' 162
    WM_NCLBUTTONDBLCLK = &HA3   ' 163
    WM_NCRBUTTONDOWN = &HA4 ' 164
    WM_NCRBUTTONUP = &HA5   ' 165
    WM_NCRBUTTONDBLCLK = &HA6   ' 166
    WM_NCMBUTTONDOWN = &HA7 ' 167
    WM_NCMBUTTONUP = &HA8   ' 168
    WM_NCMBUTTONDBLCLK = &HA9   ' 169
    WM_KEYDOWN = &H100  ' 256 - Also repeats when key held down
    WM_KEYUP = &H101    ' 257
    WM_CHAR = &H102 ' 258
    WM_DEADCHAR = &H103 ' 259
    WM_SYSKEYDOWN = &H104   ' 260
    WM_SYSKEYUP = &H105 ' 261
    WM_SYSCHAR = &H106  ' 262
    WM_SYSDEADCHAR = &H107  ' 263
    WM_CONVERTREQUESTEX = &H108 ' 264
    WM_IME_STARTCOMPOSITION = &H10D ' 269
    WM_IME_ENDCOMPOSITION = &H10E   ' 270
    WM_IME_COMPOSITION = &H10F  ' 271
    WM_IME_KEYLAST = &H10F  ' 271
    WM_INITDIALOG = &H110   ' 272
    WM_COMMAND = &H111  ' 273
    WM_SYSCOMMAND = &H112   ' 274
    WM_TIMER = &H113    ' 275
    WM_HSCROLL = &H114  ' 276
    WM_VSCROLL = &H115  ' 277
    WM_INITMENU = &H116 ' 278
    WM_INITMENUPOPUP = &H117    ' 279
    WM_MENUSELECT = &H11F   ' 287
    WM_MENUCHAR = &H120 ' 288
    WM_ENTERIDLE = &H121    ' 289
    WM_MENURBUTTONUP = &H122    ' 290
    WM_MENUDRAG = &H123 ' 291
    WM_MENUGETOBJECT = &H124    ' 292
    WM_MENUCOMMAND = &H126  ' 294
    WM_CTLCOLORMSGBOX = &H132   ' 306
    WM_CTLCOLOREDIT = &H133 ' 307
    WM_CTLCOLORLISTBOX = &H134  ' 308
    WM_CTLCOLORBTN = &H135  ' 309
    WM_CTLCOLORDLG = &H136  ' 310
    WM_CTLCOLORSCROLLBAR = &H137    ' 311
    WM_CTLCOLORSTATIC = &H138   ' 312
    WM_MOUSEMOVE = &H200    ' 512
    WM_LBUTTONDOWN = &H201  ' 513
    WM_LBUTTONUP = &H202    ' 514
    WM_LBUTTONDBLCLK = &H203    ' 515
    WM_RBUTTONDOWN = &H204  ' 516
    WM_RBUTTONUP = &H205    ' 517
    WM_RBUTTONDBLCLK = &H206    ' 518
    WM_MBUTTONDOWN = &H207  ' 519
    WM_MBUTTONUP = &H208    ' 520
    WM_MBUTTONDBLCLK = &H209    ' 521
    WM_MOUSEWHEEL = &H20A   ' 522
    WM_PARENTNOTIFY = &H210 ' 528
    WM_ENTERMENULOOP = &H211    ' 529
    WM_EXITMENULOOP = &H212 ' 530
    WM_NEXTMENU = &H213 ' 531
    WM_SIZING = &H214   ' 532
    WM_CAPTURECHANGED_R = &H215 ' 533
    WM_MOVING = &H216   ' 534
    WM_POWERBROADCAST = &H218   ' 536
    WM_DEVICECHANGE = &H219 ' 537
    WM_MDICREATE = &H220    ' 544
    WM_MDIDESTROY = &H221   ' 545
    WM_MDIACTIVATE = &H222  ' 546
    WM_MDIRESTORE = &H223   ' 547
    WM_MDINEXT = &H224  ' 548
    WM_MDIMAXIMIZE = &H225  ' 549
    WM_MDITILE = &H226  ' 550
    WM_MDICASCADE = &H227   ' 551
    WM_MDIICONARRANGE = &H228   ' 552
    WM_MDIGETACTIVE = &H229 ' 553
    WM_MDISETMENU = &H230   ' 560
    WM_ENTERSIZEMOVE = &H231    ' 561
    WM_EXITSIZEMOVE = &H232 ' 562
    WM_DROPFILES = &H233    ' 563
    WM_MDIREFRESHMENU = &H234   ' 564
    WM_IME_SETCONTEXT = &H281   ' 641
    WM_IME_NOTIFY = &H282   ' 642
    WM_IME_CONTROL = &H283  ' 643
    WM_IME_COMPOSITIONFULL = &H284  ' 644
    WM_IME_SELECT = &H285   ' 645
    WM_IME_CHAR = &H286 ' 646
    WM_IME_KEYDOWN = &H290  ' 656
    WM_IME_KEYUP = &H291    ' 657
    WM_MOUSEHOVER = &H2A1   ' 673
    WM_MOUSELEAVE = &H2A3   ' 675
    WM_CUT = &H300  ' 768
    WM_COPY = &H301 ' 769
    WM_PASTE = &H302    ' 770
    WM_CLEAR = &H303    ' 771
    WM_UNDO = &H304 ' 772
    WM_RENDERFORMAT = &H305 ' 773
    WM_RENDERALLFORMATS = &H306 ' 774
    WM_DESTROYCLIPBOARD = &H307 ' 775
    WM_DRAWCLIPBOARD = &H308    ' 776
    WM_PAINTCLIPBOARD = &H309   ' 777
    WM_VSCROLLCLIPBOARD = &H30A ' 778
    WM_SIZECLIPBOARD = &H30B    ' 779
    WM_ASKCBFORMATNAME = &H30C  ' 780
    WM_CHANGECBCHAIN = &H30D    ' 781
    WM_HSCROLLCLIPBOARD = &H30E ' 782
    WM_QUERYNEWPALETTE = &H30F  ' 783
    WM_PALETTEISCHANGING = &H310    ' 784
    WM_PALETTECHANGED = &H311   ' 785
    WM_HOTKEY = &H312   ' 786
    WM_PRINT = &H317    ' 791
    WM_PRINTCLIENT = &H318  ' 792
    WM_APPCOMMAND = &H319   ' 793
    WM_PENWINFIRST = &H380  ' 896
    WM_PENWINLAST = &H38F   ' 911
    WM_DDE_FIRST = &H3E0    ' 992
    WM_DDE_INITIATE = &H3E0 ' 992
    WM_DDE_TERMINATE = (&H3E0 + 1)  ' 993
    WM_DDE_ADVISE = (&H3E0 + 2) ' 994
    WM_DDE_UNADVISE = (&H3E0 + 3)   ' 995
    WM_DDE_ACK = (&H3E0 + 4)    ' 996
    WM_DDE_DATA = (&H3E0 + 5)   ' 997
    WM_DDE_REQUEST = (&H3E0 + 6)    ' 998
    WM_DDE_POKE = (&H3E0 + 7)   ' 999
    WM_DDE_EXECUTE = (&H3E0 + 8)    ' 1000
    WM_DDE_LAST = (&H3E0 + 8)   ' 1000
    WM_PSD_PAGESETUPDLG = (&H400)   ' 1024
    WM_USER = &H400 ' 1024
    WM_CHOOSEFONT_GETLOGFONT = &H401    ' 1025
    WM_PSD_FULLPAGERECT = (&H400 + 1)   ' 1025
    WM_PSD_MINMARGINRECT = (&H400 + 2)  ' 1026
    WM_PSD_MARGINRECT = (&H400 + 3) ' 1027
    WM_PSD_GREEKTEXTRECT = (&H400 + 4)  ' 1028
    WM_PSD_ENVSTAMPRECT = (&H400 + 5)   ' 1029
    WM_PSD_YAFULLPAGERECT = (&H400 + 6) ' 1030
    WM_CHOOSEFONT_SETLOGFONT = (&H400 + 101)    ' 1125
    WM_CHOOSEFONT_SETFLAGS = (&H400 + 102)  ' 1126
End Enum

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function MyWindowProc(ByVal hwnd As Long, ByVal wMsg As WindowMessages, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Debug.Print ("wMsg = " & CStr(wMsg))
    Select Case wMsg
        'Put the events you want to ignore/act-on on in here
        Case WM_RBUTTONDOWN
            Debug.Print ("Right button has been disabled")
        Case WM_COPY
            Debug.Print ("Ctrl-C has been disabled")
        Case WM_CUT
            Debug.Print ("Ctrl-X has been disabled")
        Case WM_PASTE
            Debug.Print ("Ctrl-V has been disabled")
        Case WM_CONTEXTMENU
            Debug.Print ("Menu Key has been disabled")
        Case Else
            MyWindowProc = CallWindowProc(GetWindowLong(hwnd, GWL_USERDATA), hwnd, wMsg, wParam, lParam)
    End Select
End Function

' Recommended Usage
'
' Private Sub Form_Load()
'     Call SubClassWnd( MyObject.hwnd )
' End Sub

Public Sub SubClassWnd(hwnd As Long)
    Call SetWindowLong(hwnd, GWL_USERDATA, SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MyWindowProc))
End Sub

' Recommended Usage
'
' Private Sub Form_Unload(Cancel As Integer)
'     Call UnSubclassWnd( MyObject.hwnd )
' End Sub

Sub UnSubclassWnd(hwnd As Long)
    Call SetWindowLong(hwnd, GWL_WNDPROC, GetWindowLong(hwnd, GWL_USERDATA))
    Call SetWindowLong(hwnd, GWL_USERDATA, 0)
End Sub
