VERSION 5.00
Begin VB.UserControl ucWizzForm 
   CanGetFocus     =   0   'False
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   570
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucWizzForm.ctx":0000
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   38
   ToolboxBitmap   =   "ucWizzForm.ctx":0C44
End
Attribute VB_Name = "ucWizzForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'   ucWizzForm.ctl
'
'
'   Start      :    06/14/2004
'   Created by :    LightTempler
'   last edit  :    10/12/2005
'
'
'   What       :    Often needed functions and extensions when working with VB forms in a fully self-contained
'                   usercontrol (thx&credits to Paul Caton! - no more bas file(s), IDE crash or anything else on subclassing).
'
'
'                   * Restrict forms size (min/max) flickerfree.
'                   * Get much more events (app activate/deactivate, resize start/end, move start/end, ... )
'                   * Save/restore form's last size/position to registry or INI file (still yet:  No line of code ... ;) )
'                   * Two more buttons in form's caption bar:
'                                   1 - Collapse to headline/restore to previous size
'                                   2 - Stay on top of all other windows
'                   * Form resizing with auto keep aspect ratio.
'                   * A simple and powerfull solution to resize/reposition controls on form when resizing the form.
'                   * Drag form full or frame only independent to control panels setting.
'                   * Center/resize form to useable desktop area with respect to all taskbars. A gap to the borders can be given.
'                   * Show form without activation.
'                   * Simple interface to receive string messages send by API SendMsg.
'                   * Some usefull properties.
'                   * And more ...
'
'                   All of this features are switchable/configurable on design time in WizzForms usercontrol
'                   property window. No line of code neccessary!
'                   Switch them on or off, single (or all together with the 'Enabled' property).
'
'   What not   :    * The usercontrol-on-a-form problem:
'                     Normaly when you put an usercontrol onto a form you got a very special problem because of a VB bug:
'                     When you create a link to your compiled VB exe and set the link property e.g. to'show minimized' than
'                     when starting your app using this link your app appears - in normal mode, NOT minimized ... :(
'                     And the same for 'show maximized' ...
'                     WizzFom fixes this for your form. Use as many other usercontrols as you like.
'                   * Here are no fancy explode-the-form-in/out effects, no form-as-a-star region functions and
'                     no transparency or hole-into-the-form-gimmicks.
'                   * Problems with older VB versions. All the stuff here should run without trouble in VB5 and VB6.
'                     No support for look of the additional buttons in forms caption bar on themed XP - I couldn't check it out,
'                     because I don't have/use Win XP. MS has changed the design/painting so maybe the 'PaintCaptionButton()'
'                     needs some additions.
'
'
'
'   Credits    :    PAUL CATON      This usercontrol uses Paul's fantastic subclassing code with "inline"-assembler
'                                   technik - v1.1.0005 20040620 - released on WWW.Planet-Source-Code.Com / VB section.
'                                   Definitions and code are reformated and integrated by me.
'
'                   BRYAN STAFFORD  He released a demo for a button in form's caption bar. I reused some of his code and
'                                   solutions.
'
'                   SIMON MORGAN    His submission on PSC shows me a working way to put a form on top of all others,
'                                   even on W2K and WinXP.
'
'                   Carles P.V.     Thx for the fastest compatible gradient sub ever in pure VB!
'
'                   Abstractvb.com  Vbspeed.com says (after good tests) its the fastest solution to split a RGB value
'                                   into components
'
'                   aboutvb.de      Article with bas to fix the VB bug, when a usercontrol is placed onto a form.
'                                   'StartUpWindowState()'
'
'
'
'
'   Specials    :   DrawTopDownGradient(), PersistentValueSave(), PersistentValueLoad(), ResizeControls()
'
'
'   Copyright   :   (C) by Light Templer. Free to use in any project. Don't build an ActiveX from and sell it as yours ...
'                   Use at your own risk - I'm definitly NOT responsible for anything ;) .
'
'   Contact     :   schwepps_bitterlemon@gmx.de         Any sensefull comments/suggestions/improvements are welcome ;)
'
'
'
'   Update 0    :   V. 1.02 First release / no updates yet
'
'

'   Maybe you like this tool and its usefull to you.
'
'                                              LiTe


Option Explicit



' *******************************
' *            EVENTS           *
' *******************************
Public Event AppActivated()
Public Event AppDeactivated()
Public Event Error(lErrNo As Long, sErrMsg As String)
Public Event FormCollapse(flgShrink As Boolean)                 ' Raised, when the additional button in forms caption bar
                                                                ' is pressed. (With flgShrink = True when form will shrink down)
                                                                ' The event is raised just BEFORE sizing, so you can change
                                                                ' values like 'Min Size' before ;) ...
Public Event FormMoveSizeStart()
Public Event FormMoveSizeEnd()
Public Event FormMoving()
Public Event FormResizing(Edge As enWMSZ)
Public Event FormSizeStatechanged(enWMSizeStateChange)          ' Form minimized, maximized, restored, other form maximized
Public Event FormStayOnTop(flgActiated As Boolean)              ' Raised, when the additional button in forms caption bar
                                                                ' is pressed. (With flgActivated = True when form will stay on top)
Public Event MouseEntersForm()
Public Event MouseLeavesForm()                                  ' Reliably. Even when mouse "jumps" from a control to outside form.
Public Event ReceivedMessage(sMessage As String)                ' Msg received from a WM_COPYDATA send with specified number (MsgHandle)
Public Event SystemColorChange(flgEffectOfThemeChange As Boolean)
Public Event ScreenResChanged(NewWidth As Long, NewHeight As Long, NewColorDepth As Long)
Public Event ThemeChange()




' *************************************
' *        PUBLIC ENUMS               *
' *************************************

' Used with property 'SaveIn'
Public Enum enSaveIn                                                            ' With public enums you run into trouble, when
    WF_SI_Registry = 1                                                          ' using different usercontrols with same enum
    WF_SI_INI = 2                                                               ' names (famous examples 'Left', 'None', 'Yes', ...)
End Enum                                                                        ' To avoid this I always try to use unique names
                                                                                ' following a naming standard.
' Used with event 'WM_SIZING'
Public Enum enWMSZ
    WF_WMSZ_LEFT = 1
    WF_WMSZ_RIGHT = 2
    WF_WMSZ_TOP = 3
    WF_WMSZ_TOPLEFT = 4
    WF_WMSZ_TOPRIGHT = 5
    WF_WMSZ_BOTTOM = 6
    WF_WMSZ_BOTTOMLEFT = 7
    WF_WMSZ_BOTTOMRIGHT = 8
End Enum

Public Enum enWMSizeStateChange
    WF_WMSSC_SIZE_RESTORED = 0
    WF_WMSSC_SIZE_MINIMIZED = 1
    WF_WMSSC_SIZE_MAXIMIZED = 2
    WF_WMSSC_SIZE_MAXSHOW = 3
    WF_WMSSC_SIZE_MAXHIDE = 4
End Enum

Public Enum enWFGradient
    WF_GR_None = 0
    WF_GR_TwoColorGradient = 1
    WF_GR_TwoColorGradPlusBlock = 2
    WF_GR_ThreeColorGradient = 3
End Enum

Public Enum enFullDrag
    WF_FD_DontChange = 0
    WF_FD_Yes = 1
    WF_FD_No = 2
End Enum

Public Enum enUnit                                          ' Most commonly used, add more if you like and make
    WF_UN_Pixels = 0                                        ' additions to the functions 'ValToUsedUnit()' and 'UsedUnitToValue()'.
    WF_UN_Twips = 1                                         ' Notice: Internaly all values are saved in 'Pixels'
End Enum



' *************************************
' *            CONSTANTS              *
' *************************************

Private Const PersistentSection = "WizzFormPersistents"     ' Section name for saving persistent values in registry or INI file
                                                            ' If SaveInINI is selected:  By default
                                                            '        App.Path + "\" + App.EXEName + ".Ini"
                                                            ' is used as path/filename for the INI file.
                                                            ' To change this behavior just goto  GetINIPathName()

Private Const TAG_DELIMITER = "|"                           ' Used in Tag Property of controls for autoresizing. Change to your
                                                            ' needs, e.g.  "~"  or  ","  or any other.

Private Const TIMER_ID_WF = 2201                            ' A random 'magic number' to identify the API soft timer in subclassing.
                                                            ' This timer starts on first WM_MOUSEMOVE to get an reliably
                                                            ' "mouse leaves form" event and is killed when mouse leaves. So we
                                                            ' don't have problems when mouse is moved from a control on form to
                                                            ' an outside position in a fast way. (WM_MOUSELEAVE fails on this!)


' Property defauts

Private Const m_def_Enabled             As Boolean = True
Private Const m_def_SavePosition        As Boolean = False
Private Const m_def_SaveSize            As Boolean = False
Private Const m_def_AutoResizeControls  As Boolean = True
Private Const m_def_AdditionalEvents    As Boolean = True
Private Const m_def_StayOnTop           As Boolean = False
Private Const m_def_KeepAspectRatio     As Boolean = False
Private Const m_def_BtnCollapse         As Boolean = False
Private Const m_def_BtnStayOnTop        As Boolean = False
Private Const m_def_FormSunken          As Boolean = False
Private Const m_def_CollapseSmallSize   As Long = 0
Private Const m_def_FormMinWidth        As Long = 0
Private Const m_def_FormMinHeight       As Long = 0
Private Const m_def_FormMaxWidth        As Long = 0
Private Const m_def_FormMaxHeight       As Long = 0
Private Const m_def_FormMaxPosX         As Long = 0
Private Const m_def_FormMaxPosY         As Long = 0
Private Const m_def_BGWidth             As Long = 0
Private Const m_def_BGColorChange       As Long = 0
Private Const m_def_BGColor             As Long = vbWhite                       ' You cannot declare 'As Enum' :( ...
Private Const m_def_SaveIn              As Long = enSaveIn.WF_SI_Registry
Private Const m_def_FullDrag            As Long = enFullDrag.WF_FD_DontChange
Private Const m_def_BackgroundGradient  As Long = enWFGradient.WF_GR_None
Private Const m_def_Unit                As Long = enUnit.WF_UN_Pixels


' Subclass stuff
Private Const ALL_MESSAGES              As Long = -1&           ' All messages added or deleted
Private Const GMEM_FIXED                As Long = 0&
Private Const GWL_WNDPROC               As Long = -4&           ' Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04                  As Long = 88&           ' Table B (before) address patch offset
Private Const PATCH_05                  As Long = 93&           ' Table B (before) entry count patch offset
Private Const PATCH_08                  As Long = 132&          ' Table A (after) address patch offset
Private Const PATCH_09                  As Long = 137&          ' Table A (after) entry count patch offset
                                                                

' API - Windows messages
Private Const WM_SIZE                   As Long = &H5&
Private Const WM_SYSCOLORCHANGE         As Long = &H15&
Private Const WM_ACTIVATEAPP            As Long = &H1C&
Private Const WM_GETMINMAXINFO          As Long = &H24&
Private Const WM_THEMECHANGED           As Long = &H31A&
Private Const WM_STYLECHANGED           As Long = &H7D&
Private Const WM_DISPLAYCHANGE          As Long = &H7E&
Private Const WM_NCPAINT                As Long = &H85&
Private Const WM_NCACTIVATE             As Long = &H86&
Private Const WM_NCLBUTTONDOWN          As Long = &HA1&
Private Const WM_NCLBUTTONDBLCLK        As Long = &HA3&
Private Const WM_TIMER                  As Long = &H113&
Private Const WM_MOUSEMOVE              As Long = &H200&
Private Const WM_LBUTTONUP              As Long = &H202&
Private Const WM_SIZING                 As Long = &H214&
Private Const WM_MOVING                 As Long = &H216&
Private Const WM_ENTERSIZEMOVE          As Long = &H231&
Private Const WM_EXITSIZEMOVE           As Long = &H232&
Private Const WM_COPYDATA               As Long = &H4A&


' wParam of WM_SIZE
Private Const SIZE_RESTORED             As Long = &H0&      ' The window has been resized, but neither the SIZE_MINIMIZED nor SIZE_MAXIMIZED value applies.
Private Const SIZE_MINIMIZED            As Long = &H1&      ' The window has been minimized
Private Const SIZE_MAXIMIZED            As Long = &H2&      ' The window has been maximized
Private Const SIZE_MAXSHOW              As Long = &H3&      ' Msg is sent to all pop-up windows when some other window has been restored to its former size.
Private Const SIZE_MAXHIDE              As Long = &H4&      ' Msg is sent to all pop-up windows when some other window is maximized.


' API - System metrics
Private Const SM_CYCAPTION              As Long = 4&
Private Const SM_CXSIZE                 As Long = 30&
Private Const SM_CYSIZE                 As Long = 31&
Private Const SM_CXEDGE                 As Long = 45&
Private Const SM_CYEDGE                 As Long = 46&
Private Const SM_CXSMSIZE               As Long = 52&
Private Const SM_CYSMSIZE               As Long = 53&


' API - Misc
Private Const SW_SHOWNOACTIVATE         As Long = 4&        ' Used with 'API_SetWindowLong()'
Private Const SWP_NOSIZE                As Long = 1&
Private Const SWP_NOMOVE                As Long = 2&
Private Const HWND_TOPMOST              As Long = -1&
Private Const HWND_NOTOPMOST            As Long = -2&
Private Const SPI_GETWORKAREA           As Long = 48&       ' Used by  'SystemParametersInfo()'
Private Const SPI_GETDRAGFULLWINDOWS    As Long = 38&
Private Const SPI_SETDRAGFULLWINDOWS    As Long = 37&
Private Const SPIF_SENDWININICHANGE     As Long = 2&
Private Const GWL_STYLE                 As Long = -16&
Private Const GWL_EXSTYLE               As Long = -20&
Private Const WS_EX_CLIENTEDGE          As Long = &H200&
Private Const WS_EX_TOOLWINDOW          As Long = &H80&
Private Const DFC_BUTTON                As Long = 4&        ' DrawFrameControl:  'Standard button'
Private Const DFCS_BUTTONPUSH           As Long = 16&       ' DrawFrameControl:  'Push button'
Private Const DFCS_PUSHED               As Long = &H200&    ' DrawFrameControl:  'Push button - pressed'
Private Const WS_THICKFRAME             As Long = &H40000
Private Const HTCAPTION                 As Long = 2&        ' Used in subclass: click in caption bar of form
Private Const API_INVALID_COLOR         As Long = -1&       ' Result from 'OLEColorToRGB()' when OleTranslateColor() fails
Private Const API_DIB_RGB_COLORS        As Long = 0&        ' Used for gradient sub

 

' *************************************
' *        PRIVATE ENUMS              *
' *************************************
Private Enum enMsgWhen
    MSG_AFTER = 1                                           ' Message calls back after the original WndProc
    MSG_BEFORE = 2                                          ' Message calls back before the original WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE          ' Message calls back before and after the original WndProc
End Enum

Private Enum enBtnSymbol                                    ' Used in 'PaintCaptionButton()'
    BS_ArrowUp = 1
    BS_ArrowDown = 2
    BS_DontStayOnTop = 3
    BS_StayOnTop = 4
End Enum



' *************************************
' *        PRIVATE TYPES              *
' *************************************
Private Type tpMvar                                         ' Here we have all usercontrol local vars together for easy
                                                            ' access with VB's intellisense.
    ' Properties
    flgEnabled                      As Boolean
    flgSubclassingActivated         As Boolean
    flgSavePosition                 As Boolean
    flgSaveSize                     As Boolean
    flgAutoResizeControls           As Boolean
    flgAdditionalEvents             As Boolean
    flgKeepAspectRatio              As Boolean
    flgResizeOnMaximize             As Boolean              ' Ensures to Resize Controls on Maximize/Restore once only
    lFormPosX                       As Long
    lFormPosY                       As Long
    lFormHeight                     As Long
    lFormWidth                      As Long
    lFormMinWidth                   As Long
    lFormMinHeight                  As Long
    lFormMaxWidth                   As Long
    lFormMaxHeight                  As Long
    lFormMaxPosX                    As Long
    lFormMaxPosY                    As Long
    lCollapseSmallSize              As Long                 ' Collapse button pressed: Resize form to this value  (default: 0)
    lCollapseOrgSize                As Long                 ' Here we save original form height to restore later
    lMsgHandle                      As Long                 ' <> 0 : Used to identify string messages send with WM_COPYDATA dedicated
                                                            ' to be handled by WizzForm.
    FullDrag                        As enFullDrag
    SaveInRegOrIni                  As enSaveIn
    BckgrndGradient                 As enWFGradient
    BGWidth                         As Long                 ' Width of gradient from left border in pixel. 0 - means: full size
                                                            ' Plz look at 'DrawBackgroundGradient()' for more infos.
    BGColorChange                   As Long                 ' Range from 0% to 100% : Marker for changing to 2nd gradient
    BGColor1                        As OLE_COLOR            ' Background gradient color 1
    BGColor2                        As OLE_COLOR            ' Background gradient color 2
    BGColor3                        As OLE_COLOR            ' Background gradient color 3
    
    ' Additional buttons in caption bar
    flgBtnCollapse                  As Boolean              ' Enabled?
    flgCollapsed                    As Boolean              ' Is form collapsed?
    flgBtnClpsPressed               As Boolean              ' Button pressed (but not released)
    
    flgBtnStayOnTop                 As Boolean
    flgStayOnTop                    As Boolean              ' Is 'stay-on-top' activated?
    flgBtnSOTPressed                As Boolean              ' Button pressed (but not release)
    
    flgFormSunken                   As Boolean              ' API Sunken Form Mode avtivated?
    
    ' Misc
    flgFormLoadDone                 As Boolean              ' Flag: Is form already loaded?
    Unit                            As enUnit               ' Unit for values used
    
End Type


' === UDTs used by API calls
Private Type tSubData                                       ' Subclass data type
  hwnd                              As Long                 ' Handle of the window being subclassed
  nAddrSub                          As Long                 ' The address of our new WndProc (aCode)
  nAddrOrig                         As Long                 ' The address of the pre-existing WndProc
  nMsgCntA                          As Long                 ' Msg after table entry count
  nMsgCntB                          As Long                 ' Msg before table entry count
  aMsgTblA()                        As Long                 ' Msg after table array
  aMsgTblB()                        As Long                 ' Msg Before table array
End Type
Private sc_aSubData()               As tSubData             ' Subclass data array
                                    
Private Type tpAPI_POINT
    X                               As Long
    Y                               As Long
End Type
    
Private Type tpAPI_RECT                                     ' NEVER ever use 'Left' or 'Right' as names in a udt!
    lLeft                           As Long                 ' You run into trouble with the Vb build-in functions for
    lTop                            As Long                 ' string/variant handling. And this strange effects and error
    lRight                          As Long                 ' messages are really hard to debug ... ;(
    lBottom                         As Long
End Type
    
Private Type tpAPI_MINMAXINFO
    Reserved                        As tpAPI_POINT          ' Reserved
    MaxSize                         As tpAPI_POINT          ' Form size when maximized
    MaxPosition                     As tpAPI_POINT          ' Position when maximized
    MinTrackSize                    As tpAPI_POINT          ' Min windows size on resizing
    MaxTrackSize                    As tpAPI_POINT          ' Max windows size on resizing
End Type

Private Type tpLOGFONT                                      ' Usd in 'BuildMarlettFont()'
    lfHeight                        As Long
    lfWidth                         As Long
    lfEscapement                    As Long
    lfOrientation                   As Long
    lfWeight                        As Long
    lfItalic                        As Byte
    lfUnderline                     As Byte
    lfStrikeOut                     As Byte
    lfCharSet                       As Byte
    lfOutPrecision                  As Byte
    lfClipPrecision                 As Byte
    lfQuality                       As Byte
    lfPitchAndFamily                As Byte
    lfFaceName                      As String * 32
End Type

Private Type tpBITMAPINFOHEADER
    biSize                          As Long
    biWidth                         As Long
    biHeight                        As Long
    biPlanes                        As Integer
    biBitCount                      As Integer
    biCompression                   As Long
    biSizeImage                     As Long
    biXPelsPerMeter                 As Long
    biYPelsPerMeter                 As Long
    biClrUsed                       As Long
    biClrImportant                  As Long
End Type

Private Type tpSTARTUPINFO                                  ' Used in 'StartUpWindowState()'
    cb                              As Long
    lpReserved                      As Long
    lpDesktop                       As Long
    lpTitle                         As Long
    dwX                             As Long
    dwY                             As Long
    dwXSize                         As Long
    dwYSize                         As Long
    dwXCountChars                   As Long
    dwYCountChars                   As Long
    dwFillAttribute                 As Long
    dwFlags                         As Long
    wShowWindow                     As Integer
    cbReserved2                     As Integer
    lpReserved2                     As Long
    hStdInput                       As Long
    hStdOutput                      As Long
    hStdError                       As Long
End Type

Private Type tpCOPYDATASTRUCT                               ' Used with WM_COPYDATA
    dwData                          As Long
    cbData                          As Long
    lpData                          As Long
End Type



' *************************************
' *        API DEFINITIONS            *
' *************************************

Private Declare Sub API_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, _
         Source As Any, _
         ByVal Length As Long)

Private Declare Function API_GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As Any, _
         ByVal lpKeyName As Any, _
         ByVal lpDefault As Any, _
         ByVal lpReturnedString As String, _
         ByVal nSize As Long, _
         ByVal lpFileName As String) As Long

Private Declare Function API_WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
         ByVal lpKeyName As Any, _
         ByVal lpString As Any, _
         ByVal lpFileName As String) As Long

Private Declare Function API_GetCursorPos Lib "user32" Alias "GetCursorPos" _
        (lpPoint As tpAPI_POINT) As Long

Private Declare Function API_ShowWindow Lib "user32" Alias "ShowWindow" _
        (ByVal hwnd As Long, _
         ByVal nCmdShow As Long) As Long

Private Declare Function API_GetWindowRect Lib "user32" Alias "GetWindowRect" _
        (ByVal hwnd As Long, _
         lpRect As tpAPI_RECT) As Long

Private Declare Function API_GetClientRect Lib "user32" Alias "GetClientRect" _
        (ByVal hwnd As Long, _
         lpRect As tpAPI_RECT) As Long

Private Declare Function API_OffsetRect Lib "user32" Alias "OffsetRect" _
        (lpRect As tpAPI_RECT, _
         ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function API_GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, _
         ByVal nIndex As Long) As Long

Private Declare Function API_SetWindowPos Lib "user32" Alias "SetWindowPos" _
        (ByVal hwnd As Long, _
         ByVal hWndInsertAfter As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal cx As Long, _
         ByVal cy As Long, _
         ByVal wFlags As Long) As Long

Private Declare Function API_MoveWindow Lib "user32" Alias "MoveWindow" _
        (ByVal hwnd As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal bRepaint As Long) As Long

Private Declare Function API_SwitchToThisWindow Lib "user32" Alias "SwitchToThisWindow" _
        (ByVal hwnd As Long, _
         ByVal hWindowState As Long) As Long

Private Declare Function API_SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
        (ByVal uAction As Long, _
         ByVal uParam As Long, _
         lpvParam As Any, _
         ByVal fuWinIni As Long) As Long

Private Declare Function API_PtInRect Lib "user32" Alias "PtInRect" _
        (lpRect As tpAPI_RECT, _
         ByVal X As Long, _
         ByVal Y As Long) As Long

Private Declare Function API_GetSystemMetrics Lib "user32" Alias "GetSystemMetrics" _
        (ByVal nIndex As Long) As Long

Private Declare Function API_IsIconic Lib "user32" Alias "IsIconic" _
        (ByVal hwnd As Long) As Long

Private Declare Function API_CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" _
        (lpLogFont As tpLOGFONT) As Long
        
Private Declare Function API_SelectObject Lib "gdi32" Alias "SelectObject" _
        (ByVal hdc As Long, _
         ByVal hObject As Long) As Long
         
Private Declare Function API_DeleteObject Lib "gdi32" Alias "DeleteObject" _
        (ByVal hObject As Long) As Long
        
Private Declare Function API_SetBkMode Lib "gdi32" Alias "SetBkMode" _
        (ByVal hdc As Long, _
         ByVal nBkMode As Long) As Long

Private Declare Function API_DrawText Lib "user32" Alias "DrawTextA" _
        (ByVal hdc As Long, _
         ByVal lpStr As String, _
         ByVal nCount As Long, _
         ByRef lpRect As tpAPI_RECT, _
         ByVal wFormat As Long) As Long
         
Private Declare Function API_GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long

Private Declare Function API_SetFocusAPI Lib "user32" Alias "SetFocus" _
        (ByVal hwnd&) As Long

Private Declare Function API_SetCapture Lib "user32" Alias "SetCapture" _
        (ByVal hwnd&) As Long

Private Declare Function API_ClientToScreen Lib "user32" Alias "ClientToScreen" _
        (ByVal hwnd As Long, _
         lpPoint As tpAPI_POINT) As Long

Private Declare Function API_OleTranslateColor Lib "oleaut32.dll" Alias "OleTranslateColor" _
        (ByVal lOLEColor As Long, _
         ByVal lHPalette As Long, _
         lColorRef As Long) As Long

Private Declare Function API_CreateSolidBrush Lib "gdi32" Alias "CreateSolidBrush" _
        (ByVal crColor As Long) As Long

Private Declare Function API_FillRect Lib "user32" Alias "FillRect" _
        (ByVal hdc As Long, _
         lpRect As tpAPI_RECT, _
         ByVal hBrush As Long) As Long

Private Declare Function API_DrawFrameControl Lib "user32" Alias "DrawFrameControl" _
        (ByVal hdc As Long, _
         lpRect As tpAPI_RECT, _
         ByVal un1 As Long, _
         ByVal un2 As Long) As Long

Private Declare Sub API_GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" _
        (lpStartupInfo As tpSTARTUPINFO)

Private Declare Function API_SetTimer Lib "user32.dll" Alias "SetTimer" _
        (ByVal hwnd As Long, _
         ByVal nIDEvent As Long, _
         ByVal uElapse As Long, _
         ByVal lpTimerFunc As Long) As Long

Private Declare Function API_KillTimer Lib "user32.dll" Alias "KillTimer" _
        (ByVal hwnd As Long, _
         ByVal nIDEvent As Long) As Long

Private Declare Function API_StretchDIBits Lib "gdi32" Alias "StretchDIBits" _
        (ByVal hdc As Long, _
         ByVal X As Long, _
         ByVal Y As Long, _
         ByVal dx As Long, _
         ByVal dy As Long, _
         ByVal SrcX As Long, _
         ByVal SrcY As Long, _
         ByVal wSrcWidth As Long, _
         ByVal wSrcHeight As Long, _
         lpBits As Any, _
         lpBitsInfo As tpBITMAPINFOHEADER, _
         ByVal wUsage As Long, _
         ByVal dwRop As Long) As Long



' Subclass part
Private Declare Function API_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" _
        (ByVal lpModuleName As String) As Long

Private Declare Function API_GetProcAddress Lib "kernel32" Alias "GetProcAddress" _
        (ByVal hModule As Long, _
         ByVal lpProcName As String) As Long

Private Declare Function GlobalAlloc Lib "kernel32" _
        (ByVal wFlags As Long, _
         ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" _
        (ByVal hMem As Long) As Long

Private Declare Function API_SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, _
         ByVal nIndex As Long, _
         ByVal dwNewLong As Long) As Long

Private Declare Function API_GetWindowDC Lib "user32" Alias "GetWindowDC" _
        (ByVal hwnd As Long) As Long




' *************************************
' *            PRIVATES               *
' *************************************
Private Mvar                    As tpMvar                       ' All local needed vars in an udt for easy access.
Private WithEvents frmParent    As Form                         ' Reference to the form Wizzform is placed on.
Attribute frmParent.VB_VarHelpID = -1
'
'
'





' *************************************
' *         PUBLIC FUNCTIONS          *
' *************************************


' ==========================================================================================================
' = Subclass handler - MUST be the first Public routine in this file. That includes public properties also =
' ==========================================================================================================
Public Sub zSubclass_Proc(ByVal flgBefore As Boolean, _
                            ByRef flgHandled As Boolean, _
                            ByRef lReturn As Long, _
                            ByRef lHwnd As Long, _
                            ByRef uMsg As Long, _
                            ByRef wParam As Long, _
                            ByRef lParam As Long)
    ' Parameters:
    '                       flgBefore  - Indicates whether the the message is being processed before or after the
    '                                    default handler - only really needed if a message is set to callback both before & after.
    '                       flgHandled - Set this variable to True in a 'before' callback to prevent the message being
    '                                    subsequently processed by the default handler... and if set, an 'after' callback
    '                       lReturn    - Set this variable as per your intentions and requirements, see the MSDN documentation
    '                                    for each individual message value.
    '                       lHwnd      - The window handle
    '                       uMsg       - The message number
    '                       wParam     - Message related data
    '                       lParam     - Message related data

    
    Static flgThemeChange       As Boolean
    Static flgMouseOverForm     As Boolean
    Static lAPITimerHandle      As Long                     ' ALLAPI.NET description of SetTimer result seems to be wrong.
                                                            ' Imho (checked on Win98) the result is the same as the set
                                                            ' times ID ... Just to be sure on other OS I use this var for KillTimer.
    Static DragModeStd          As enFullDrag
    Static hPrevCapture         As Long
    Static lMousePressedOver    As Long                     ' 0 - nothing, 1 - collapse button, 2 - stay-on-top button
    
    Dim MMInfo                  As tpAPI_MINMAXINFO
    Dim CurPosition             As tpAPI_POINT
    Dim FormPosAndSizeOld       As tpAPI_RECT
    Dim FormPosAndSizeNew       As tpAPI_RECT
    Dim RectBtn                 As tpAPI_RECT
    Dim RectForm                As tpAPI_RECT
    Dim flgWasADblClk           As Boolean
    Dim lDelta                  As Long                     ' Used on WM_SIZING for diffs between pixel positions
    Dim CopyData                As tpCOPYDATASTRUCT         ' Used with WM_COPYDATA
    Dim bytArrBufffer()         As Byte                     ' Used with WM_COPYDATA
    Dim sMessage                As String                   ' Used with WM_COPYDATA (Holds the message we got)
    
    
    With Mvar
        
        If uMsg = WM_NCLBUTTONDBLCLK Then                   ' To simply avoid this often seen "strange" effect with a
            uMsg = WM_NCLBUTTONDOWN                         ' second click to a button ... ;)
            flgWasADblClk = True                            ' Needed when not clicked to an additional button
        End If
            
        Select Case uMsg
            
            ' === EVENT:  Activate form
            Case WM_NCACTIVATE
                    
                    ' Additional buttons into caption bar (Init: Buttons not pressed)
                    With Mvar
                        .flgBtnClpsPressed = False
                        .flgBtnSOTPressed = False
                        DrawAdditionalButtonsIntoFormCaptionbar
                    End With
                    
                    
            
            ' === EVENT:  Application is being activated/deactivated
            Case WM_ACTIVATEAPP
            
                    ' *** NOTICE: THIS WINDOWS MESSAGE WILL NOT BE SEND WHEN RUNNING IN VBs IDE ! ***
                    ' *** To test any changes here you really have to  build an .exe ...          ***
                    
                    If .flgAdditionalEvents = True And flgBefore = True Then
                        If wParam = 0 Then
                            RaiseEvent AppDeactivated
                        Else
                            RaiseEvent AppActivated
                        End If
                    End If
                    
                    ' Full dragging or not
                    If wParam = 0 And flgBefore = False Then
                        SetDraggingMode DragModeStd                     ' On leaving app
                    ElseIf wParam = 1 And flgBefore = True Then
                        DragModeStd = SetDraggingMode(.FullDrag)        ' On entering app
                    End If
                    
                    
                    
            ' === EVENT:  Screen resolution or color depth has changed
            Case WM_DISPLAYCHANGE
                    If .flgAdditionalEvents = True Then
                        RaiseEvent ScreenResChanged(LoWord(lParam), HiWord(lParam), wParam)
                    End If
                    
                    DrawAdditionalButtonsIntoFormCaptionbar
                    
                    
            
            ' === EVENT:  Change in windows system colors
            Case WM_SYSCOLORCHANGE
                    ' Hint: See 'WM_THEMECHANGED' above for details
                    If flgThemeChange = True Then
                        flgThemeChange = False
                        RaiseEvent SystemColorChange(True)
                    Else
                        RaiseEvent SystemColorChange(False)
                    End If
                    
                    
                    
            ' === EVENT:  Change in Windows desktop theme
            Case WM_THEMECHANGED
                ' Comment from Paul Caton:
                '   Theme changes are almost bound to change the system colors, the theme change message comes first,
                '   therefore I'm setting a flag so that when the WM_SYSCOLORCHANGED message comes microseconds after
                '   that we don't miss the theme change message in the status bar.
                flgThemeChange = True
                RaiseEvent ThemeChange
                                            
    
    
            ' === FUNCTION:  Limit windows min/max size
            Case WM_GETMINMAXINFO
                    
                    If .lFormMinHeight + .lFormMinWidth + .lFormMaxHeight + .lFormMaxWidth > 0 Then     ' Not all are zero?
                        
                        ' Copy current MinMax infos into udt
                        API_RtlMoveMemory MMInfo, ByVal lParam, Len(MMInfo)
                        
                        ' Change to our wanted values
                        If .lFormMinWidth Then MMInfo.MinTrackSize.X = .lFormMinWidth
                        If .lFormMinHeight Then MMInfo.MinTrackSize.Y = .lFormMinHeight
                        If .lFormMaxWidth Then MMInfo.MaxTrackSize.X = .lFormMaxWidth
                        If .lFormMaxHeight Then MMInfo.MaxTrackSize.Y = .lFormMaxHeight
                        If .lFormMaxPosX Then MMInfo.MaxPosition.X = .lFormMaxPosX
                        If .lFormMaxPosY Then MMInfo.MaxPosition.Y = .lFormMaxPosY
                        
                        ' Send it back
                        API_RtlMoveMemory ByVal lParam, MMInfo, Len(MMInfo)
                    End If
            
            
            
            ' === EVENT:  Start form moving or sizing
            Case WM_ENTERSIZEMOVE
                    If .flgAdditionalEvents = True Then
                        RaiseEvent FormMoveSizeStart
                    End If
                                        
                                        
            
            ' === EVENT:  End form moving or sizing
            Case WM_EXITSIZEMOVE
                    If .flgAdditionalEvents = True Then
                        RaiseEvent FormMoveSizeEnd
                    End If
                    
            
            
            ' === EVENT:  Form is being moved
            Case WM_MOVING
                    If .flgAdditionalEvents = True Then
                        RaiseEvent FormMoving
                    End If
                    
            
            
            ' === EVENT:  Form is being sized
            Case WM_SIZING
            
                    ' New event
                    If .flgAdditionalEvents = True Then
                        RaiseEvent FormResizing(wParam)
                    End If
                    
                    ' Resize controls
                    If .flgAutoResizeControls = True Then
                        ResizeControls
                    End If
                    
                    ' Form keeps aspect ratio height/width
                    If .flgKeepAspectRatio = True Then
                                            
                        ' Get old form's size
                        API_GetWindowRect frmParent.hwnd, FormPosAndSizeOld
                        ' Get new form's size
                        API_RtlMoveMemory FormPosAndSizeNew, ByVal lParam, LenB(FormPosAndSizeNew)
                        
                        ' Modify new size to keep aspect ratio
                        With FormPosAndSizeNew
                            Select Case wParam
                                
                                Case WF_WMSZ_LEFT
                                        lDelta = .lLeft - FormPosAndSizeOld.lLeft
                                        .lTop = .lTop + lDelta
                                
                                Case WF_WMSZ_RIGHT
                                        lDelta = .lRight - FormPosAndSizeOld.lRight
                                        .lBottom = .lBottom + lDelta
                            
                                Case WF_WMSZ_TOP
                                        lDelta = .lTop - FormPosAndSizeOld.lTop
                                        .lLeft = .lLeft + lDelta
                                
                                Case WF_WMSZ_BOTTOM
                                        lDelta = .lBottom - FormPosAndSizeOld.lBottom
                                        .lRight = .lRight + lDelta
                                                        
                                Case WF_WMSZ_TOPLEFT
                                        lDelta = .lLeft - FormPosAndSizeOld.lLeft
                                        .lTop = FormPosAndSizeOld.lTop + lDelta
                                        
                                Case WF_WMSZ_BOTTOMRIGHT
                                        lDelta = .lRight - FormPosAndSizeOld.lRight
                                        .lBottom = FormPosAndSizeOld.lBottom + lDelta
                                
                                Case enWMSZ.WF_WMSZ_BOTTOMLEFT
                                        lDelta = .lLeft - FormPosAndSizeOld.lLeft
                                        .lBottom = FormPosAndSizeOld.lBottom - lDelta
                                        
                                Case WF_WMSZ_TOPRIGHT
                                        lDelta = .lRight - FormPosAndSizeOld.lRight
                                        .lTop = FormPosAndSizeOld.lTop - lDelta
                                        
                            End Select
                        End With
                        ' Put new form's size
                        API_RtlMoveMemory ByVal lParam, FormPosAndSizeNew, LenB(FormPosAndSizeNew)
                    End If
                    
                    ' New buttons in forms title bar
                    DrawAdditionalButtonsIntoFormCaptionbar
                    
                    
                    
            
            ' === EVENT:  Mouse is moved over form
            Case WM_MOUSEMOVE
            
                    ' 1 - Additional buttons - handling the situation: Moving the mouse w pressed button from addtional buttons
                    With Mvar
                        If lMousePressedOver > 0 Then
                            API_GetWindowRect frmParent.hwnd, RectForm
                        
                            If lMousePressedOver = 1 Then       ' Collapse button
                                GetButtonRect frmParent.hwnd, RectBtn, 4
                                GetScreenPoint frmParent.hwnd, lParam, CurPosition
                                If API_PtInRect(RectBtn, CurPosition.X - RectForm.lLeft, CurPosition.Y - RectForm.lTop) Then
                                    If .flgBtnClpsPressed <> True Then
                                        .flgBtnClpsPressed = True
                                        DrawAdditionalButtonsIntoFormCaptionbar
                                    End If
                                Else
                                    If .flgBtnClpsPressed <> False Then
                                        .flgBtnClpsPressed = False
                                        DrawAdditionalButtonsIntoFormCaptionbar
                                    End If
                                End If
                            End If
                                
                            If lMousePressedOver = 2 Then       ' Stay-on-top button
                                GetButtonRect frmParent.hwnd, RectBtn, 5
                                GetScreenPoint frmParent.hwnd, lParam, CurPosition
                                If API_PtInRect(RectBtn, CurPosition.X - RectForm.lLeft, CurPosition.Y - RectForm.lTop) Then
                                    If .flgBtnSOTPressed <> True Then
                                        .flgBtnSOTPressed = True
                                        DrawAdditionalButtonsIntoFormCaptionbar
                                    End If
                                Else
                                    If .flgBtnSOTPressed <> False Then
                                        .flgBtnSOTPressed = False
                                        DrawAdditionalButtonsIntoFormCaptionbar
                                    End If
                                End If
                            End If
                        End If
                    End With
                        
                        
                    ' 2 - Additional events
                    If .flgAdditionalEvents = True And flgMouseOverForm = False Then
                        RaiseEvent MouseEntersForm
                        flgMouseOverForm = True
                        
                        ' Here we start tracking the mouse postion by an API soft timer. Thats afaik the only sure way to
                        ' get the "mouse leaves form" event. WM_MOUSE_LEAVE is fine for non containers only. When mouse
                        ' "jumps" fast from a control to outside the form, you won't get an WM_MOUSE_LEAVE ...
                        lAPITimerHandle = API_SetTimer(frmParent.hwnd, TIMER_ID_WF, 300&, 0&)
                    End If
                                
                                
                        
            ' === EVENT:  Soft timer, startet in WM_MOUSEMOVE (look two lines above ;) )
            Case WM_TIMER
                    API_GetCursorPos CurPosition
                    API_GetWindowRect frmParent.hwnd, RectForm
                    
                    ' Mouse not over form anymore?
                    If API_PtInRect(RectForm, CurPosition.X, CurPosition.Y) = 0 Then
                        
                        ' Stop timer / mouse position tracking
                        API_KillTimer frmParent.hwnd, lAPITimerHandle
                        
                        If .flgAdditionalEvents = True Then     ' still wanted?
                            RaiseEvent MouseLeavesForm
                        End If
                        
                        ' Set 'Ready to activate again'
                        flgMouseOverForm = False
                    End If
                    
                    
                    
            ' === EVENT:  Repaint window / window style changed
            Case WM_NCPAINT, WM_STYLECHANGED
                    DrawAdditionalButtonsIntoFormCaptionbar
                    DrawBackgroundGradient
                    
                    
                    
            ' === EVENT:  Left mouse button pressed
            Case WM_NCLBUTTONDOWN
                    ' Click on additional buttons in caption bar?
                    If wParam = HTCAPTION Then
                        API_GetWindowRect frmParent.hwnd, RectForm
                        With Mvar
                            
                            ' 1 - Button 'Collapse'
                            If .flgBtnCollapse = True Then
                                GetButtonRect frmParent.hwnd, RectBtn, 4
                                If API_PtInRect(RectBtn, LoWord(lParam) - RectForm.lLeft, HiWord(lParam) - RectForm.lTop) Then
                                    If API_GetActiveWindow() <> frmParent.hwnd Then
                                        API_SetFocusAPI frmParent.hwnd
                                    End If
                                    .flgBtnClpsPressed = True
                                    lMousePressedOver = 1
                                    hPrevCapture = API_SetCapture(frmParent.hwnd)
                                    DrawAdditionalButtonsIntoFormCaptionbar
                                    flgHandled = True
                                End If
                            End If
                            
                            ' 2 - Button 'Stay on top'
                            If .flgBtnStayOnTop = True Then
                                GetButtonRect frmParent.hwnd, RectBtn, 5
                                If API_PtInRect(RectBtn, LoWord(lParam) - RectForm.lLeft, HiWord(lParam) - RectForm.lTop) Then
                                    If API_GetActiveWindow() <> frmParent.hwnd Then
                                        API_SetFocusAPI frmParent.hwnd
                                    End If
                                    .flgBtnSOTPressed = True
                                    lMousePressedOver = 2
                                    hPrevCapture = API_SetCapture(frmParent.hwnd)
                                    DrawAdditionalButtonsIntoFormCaptionbar
                                    flgHandled = True
                                End If
                            End If
                             
                            If flgHandled = False And flgWasADblClk = True Then         ' Not over a additional button
                                uMsg = WM_NCLBUTTONDBLCLK                               ' we need the double click msg
                                flgWasADblClk = False                                   ' on the caption bar of the form
                            End If
                        End With
                    End If
            
            
            
            ' === EVENT:  Left mouse button released
            Case WM_LBUTTONUP
                    If lMousePressedOver > 0 Then
                        API_GetWindowRect frmParent.hwnd, RectForm
                        With Mvar
                            
                            ' 1 - Button 'Collapse'
                            If .flgBtnClpsPressed = True Then
                                .flgBtnClpsPressed = False
                                
                                ' Over button?
                                GetButtonRect frmParent.hwnd, RectBtn, 4
                                GetScreenPoint frmParent.hwnd, lParam, CurPosition
                                If API_PtInRect(RectBtn, CurPosition.X - RectForm.lLeft, CurPosition.Y - RectForm.lTop) Then
                                    
                                    ' === Here is finally the action ;)
                                    .flgCollapsed = Not (.flgCollapsed)
                                    CollapseForm
                                    
                                End If
                                DrawAdditionalButtonsIntoFormCaptionbar
                                flgHandled = True
                                API_SetCapture hPrevCapture
                            End If
                            
                            ' 2 - Button 'Stay on top'
                            If .flgBtnSOTPressed = True Then
                                .flgBtnSOTPressed = False
                                
                                ' Over button?
                                GetButtonRect frmParent.hwnd, RectBtn, 5
                                GetScreenPoint frmParent.hwnd, lParam, CurPosition
                                If API_PtInRect(RectBtn, CurPosition.X - RectForm.lLeft, CurPosition.Y - RectForm.lTop) Then
                                
                                    ' === Here is finally the action ;)
                                    .flgStayOnTop = Not (.flgStayOnTop)
                                    SetFormStayOnTop .flgStayOnTop
                                    
                                End If
                                DrawAdditionalButtonsIntoFormCaptionbar
                                flgHandled = True
                                API_SetCapture hPrevCapture
                            End If
                        End With
                        
                        lMousePressedOver = 0
                    End If
                    
                    
            ' === EVENT:  The WM_SIZE message is sent to a window after its size has changed.
            Case WM_SIZE
                    
                    
                    Select Case wParam
                        
                        Case SIZE_MINIMIZED
                                If .flgAdditionalEvents = True Then
                                    RaiseEvent FormSizeStatechanged(WF_WMSSC_SIZE_MINIMIZED)
                                End If
                                   
                        Case SIZE_MAXIMIZED
                                If .flgAutoResizeControls = True Then
                                    .flgResizeOnMaximize = True
                                    ResizeControls
                                End If
                                If .flgAdditionalEvents = True Then
                                    RaiseEvent FormSizeStatechanged(WF_WMSSC_SIZE_MAXIMIZED)
                                End If
                        
                        Case SIZE_RESTORED
                                If .flgAutoResizeControls = True And .flgResizeOnMaximize = True Then
                                    ResizeControls
                                    .flgResizeOnMaximize = False
                                End If
                                If .flgAdditionalEvents = True Then
                                    RaiseEvent FormSizeStatechanged(WF_WMSSC_SIZE_MINIMIZED)
                                End If
                        
                        Case SIZE_MAXHIDE
                                If .flgAdditionalEvents = True Then
                                    RaiseEvent FormSizeStatechanged(WF_WMSSC_SIZE_MAXHIDE)
                                End If
                                
                        Case SIZE_MAXSHOW
                                If .flgAdditionalEvents = True Then
                                    RaiseEvent FormSizeStatechanged(WF_WMSSC_SIZE_MAXSHOW)
                                End If
                    
                    End Select
        
            ' === EVENT:  Got message from other app
            Case WM_COPYDATA
                    If .lMsgHandle <> 0 Then
                        API_RtlMoveMemory CopyData, ByVal lParam, Len(CopyData)
                        With CopyData
                            ReDim bytArrBufffer(1 To .cbData)
                            API_RtlMoveMemory bytArrBufffer(1), ByVal .lpData, .cbData
                            If .dwData = Mvar.lMsgHandle Then       ' Check for matching received number with wanted number
                                sMessage = StrConv(bytArrBufffer, vbUnicode)
                                sMessage = Left$(sMessage, InStr(1, sMessage, Chr$(0)) - 1)
                                
                                RaiseEvent ReceivedMessage(sMessage)
                            End If
                        End With
                    End If
        
        
        
        End Select
    End With

End Sub


Public Sub Show_About()
Attribute Show_About.VB_Description = "Gives some information to WizzForm."
Attribute Show_About.VB_UserMemId = -552
    ' Show some information

    MsgBox "      -= WizzForm =- " + vbCrLf + vbCrLf + _
            "     Extensions to a VB form" + vbCrLf + _
            "by Light Templer - June 2004 to October 2005" + vbCrLf + vbCrLf + _
            "Base for most of the functions is" + vbCrLf + _
            "Paul Catons great uc subclassing.", _
            vbInformation, "  About  - WizzForm 1.02 - "

End Sub


Public Sub ShowFormWithoutActivation()
    ' This function shows the form, but doesn't change the focus to it (doesn't selects the form).
    
    If Not frmParent Is Nothing Then
        API_ShowWindow frmParent.hwnd, SW_SHOWNOACTIVATE
    End If

End Sub


Public Sub PutFormOnTop()
    ' PUT form on top of all windows, even on W2K ... (Attention: This is NOT:  'Form will stay on top' !)
    '
    ' Hint:     All other (simple) API call solutions doesn't work on W2K or higher.
    '           Taskbar entry is blinking, but form is not on top. This
    '           undocumented call is tested on W98, NT4, W2K and XP and does the job.
    '
    '           Thx to Simon Morgan for his submission on PSC to this API call
    
    If Not frmParent Is Nothing Then
        API_SwitchToThisWindow frmParent.hwnd, vbNormalFocus
    End If
    
End Sub


Public Sub CenterFormInWorkArea(Optional ByVal lGapToEdges As Long = -1)
    ' Get the useable area of the desktop (no taskbar, officebar, ...) and
    ' center ther form within. Specifying a gap to screen borders means:
    ' 1st - Resize form. 2nd Now center form
    '
    ' Parameter:
    '           lGapToEdges:     = -1   Don't change the size of the form.
    '                           >=  0   Resize form to keep this distance to the
    '                                   useable desktop borders. Selected UNIT is
    '                                   is used for this!
    
    
    Const ForceRefresh As Long = 1&
    
    Dim RectWorkArea    As tpAPI_RECT
    Dim RectWindow      As tpAPI_RECT
    Dim RectWindowNew   As tpAPI_RECT   ' lRight/lBottom abused for width/height
    
    ' No form, no fun ...
    If frmParent Is Nothing Then
        
        Exit Sub
    End If
    
    ' Some error checking ...
    If lGapToEdges < -1 Or _
            frmParent.WindowState <> vbNormal Or _
            (frmParent.BorderStyle <> vbSizable And frmParent.BorderStyle <> vbSizableToolWindow) Then
    
        Exit Sub
    End If
    
    ' Get useable area
    API_SystemParametersInfo SPI_GETWORKAREA, 0, RectWorkArea, 0
    API_GetWindowRect frmParent.hwnd, RectWindow
    
    If lGapToEdges > (RectWorkArea.lRight - RectWorkArea.lLeft) + 30 Then       ' Distance to borders too large!
    
        Exit Sub
    End If
    
    With RectWindowNew
        If lGapToEdges = -1 Then
            
            ' Just center form
            .lRight = RectWindow.lRight - RectWindow.lLeft      ' Width  - No change
            .lBottom = RectWindow.lBottom - RectWindow.lTop     ' Height - No change
            
            .lLeft = RectWorkArea.lLeft + ((RectWorkArea.lRight - RectWorkArea.lLeft) / 2 - (.lRight / 2))
            .lTop = RectWorkArea.lTop + ((RectWorkArea.lBottom - RectWorkArea.lTop) / 2 - (.lBottom / 2))
        Else
            
            lGapToEdges = UsedUnitToValue(lGapToEdges)
            
            ' Resize / center form
            .lLeft = RectWorkArea.lLeft + lGapToEdges
            .lTop = RectWorkArea.lTop + lGapToEdges
        
            .lRight = (RectWorkArea.lRight - RectWorkArea.lLeft) - (2 * lGapToEdges)
            .lBottom = (RectWorkArea.lBottom - RectWorkArea.lTop) - (2 * lGapToEdges)
        End If
        
        ' Move/resize to new values
        API_MoveWindow frmParent.hwnd, .lLeft, .lTop, .lRight, .lBottom, ForceRefresh
    End With
    DoEvents
                
End Sub

Public Function ResetPersistentSavedWindowPosition()
    ' Used to reset any further (by WizzForm) persistent saved windows size and position
    ' Maybe usefull in a config dialog behind a Reset-button or with command line start switch
    ' Attention:  The flag 'Save in registry' / 'Save in INI file' must match ... ;-)
    
    With frmParent
        PersistentValueDelete .Name + "-" + "FormPosX"
        PersistentValueDelete .Name + "-" + "FormPosY"
        PersistentValueDelete .Name + "-" + "FormWidth"
        PersistentValueDelete .Name + "-" + "FormHeight"
    End With
    
End Function



Public Function HiWord(ByVal dwValue As Long) As Integer
    ' Returns the high 16-bit integer from a 32-bit long integer
    
    API_RtlMoveMemory HiWord, ByVal VarPtr(dwValue) + 2, 2&
    
End Function

Public Function LoWord(ByVal dwValue As Long) As Integer
    ' Returns the low 16-bit integer from a 32-bit long integer
    
    API_RtlMoveMemory LoWord, dwValue, 2&
    
End Function





' *************************************
' *         PRIVATE FUNCTIONS         *
' *************************************

' === USERCONTROL FUNCTIONS ===

Private Sub UserControl_InitProperties()
    
    ' Set default properties (When WizzForm uc is placed on the form the first time)
    With Mvar
        .flgEnabled = m_def_Enabled
        .flgSavePosition = m_def_SavePosition
        .flgSaveSize = m_def_SaveSize
        .flgAdditionalEvents = m_def_AdditionalEvents
        .flgAutoResizeControls = m_def_AutoResizeControls
        .flgStayOnTop = m_def_StayOnTop
        .flgKeepAspectRatio = m_def_KeepAspectRatio
        .FullDrag = m_def_FullDrag
        .flgBtnCollapse = m_def_BtnCollapse
        .flgBtnStayOnTop = m_def_BtnStayOnTop
        .flgFormSunken = m_def_FormSunken
        .lCollapseSmallSize = m_def_CollapseSmallSize
        .lFormMinWidth = m_def_FormMinWidth
        .lFormMinHeight = m_def_FormMinHeight
        .lFormMaxWidth = m_def_FormMaxWidth
        .lFormMaxHeight = m_def_FormMaxHeight
        .lFormMaxPosX = m_def_FormMaxPosX
        .lFormMaxPosY = m_def_FormMaxPosY
        .SaveInRegOrIni = m_def_SaveIn
        .BckgrndGradient = m_def_BackgroundGradient
        .BGWidth = m_def_BGWidth
        .BGColorChange = m_def_BGColorChange
        .BGColor1 = m_def_BGColor
        .BGColor2 = m_def_BGColor
        .BGColor3 = m_def_BGColor
        
        On Local Error Resume Next
        .Unit = IIf(Ambient.ScaleUnits = "Pixel", enUnit.WF_UN_Pixels, enUnit.WF_UN_Twips)      ' Needs more code for handling
    End With                                                                                    ' different units like 'Inch' ...
    
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Local Error Resume Next
    
    With Mvar
        .flgEnabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
        .flgSavePosition = PropBag.ReadProperty("SavePosition", m_def_SavePosition)
        .flgSaveSize = PropBag.ReadProperty("SaveSize", m_def_SaveSize)
        .flgAdditionalEvents = PropBag.ReadProperty("AdditionalEvents", m_def_AdditionalEvents)
        .flgAutoResizeControls = PropBag.ReadProperty("AutoResizeControls", m_def_AutoResizeControls)
        .flgStayOnTop = PropBag.ReadProperty("StayOnTop", m_def_StayOnTop)
        .flgKeepAspectRatio = PropBag.ReadProperty("KeepAspectRatio", m_def_KeepAspectRatio)
        .FullDrag = PropBag.ReadProperty("FullDrag", m_def_FullDrag)
        .flgBtnCollapse = PropBag.ReadProperty("CollapseButton", m_def_BtnCollapse)
        .flgBtnStayOnTop = PropBag.ReadProperty("StayOnTopButton", m_def_BtnStayOnTop)
        .flgFormSunken = PropBag.ReadProperty("FormSunken", m_def_FormSunken)
        .lCollapseSmallSize = PropBag.ReadProperty("CollapseSmallSize", m_def_CollapseSmallSize)
        .lFormMinWidth = PropBag.ReadProperty("FormMinWidth", m_def_FormMinWidth)
        .lFormMinHeight = PropBag.ReadProperty("FormMinHeight", m_def_FormMinHeight)
        .lFormMaxWidth = PropBag.ReadProperty("FormMaxWidth", m_def_FormMaxWidth)
        .lFormMaxHeight = PropBag.ReadProperty("FormMaxHeight", m_def_FormMaxHeight)
        .lFormMaxPosX = PropBag.ReadProperty("FormMaxPosX", m_def_FormMaxPosX)
        .lFormMaxPosY = PropBag.ReadProperty("FormMaxPosY", m_def_FormMaxPosY)
        .lMsgHandle = PropBag.ReadProperty("MsgHandle", 0)
        .SaveInRegOrIni = PropBag.ReadProperty("SaveIn", m_def_SaveIn)
        .BckgrndGradient = PropBag.ReadProperty("BackgroundGradient", m_def_BackgroundGradient)
        .BGWidth = PropBag.ReadProperty("BGWidth", m_def_BGWidth)
        .BGColorChange = PropBag.ReadProperty("BGColorChange", m_def_BGColorChange)
        .BGColor1 = PropBag.ReadProperty("BGColor1", m_def_BGColor)
        .BGColor2 = PropBag.ReadProperty("BGColor2", m_def_BGColor)
        .BGColor3 = PropBag.ReadProperty("BGColor3", m_def_BGColor)
        .Unit = PropBag.ReadProperty("Unit", m_def_Unit)
    End With
    
    If Mvar.flgEnabled = True Then
        ActivateFunctions
    End If
        
End Sub

Private Sub UserControl_Resize()
    
   ' Const size. (WizzForms small image size)
   UserControl.Width = 480
   UserControl.Height = 480
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With Mvar
        PropBag.WriteProperty "Enabled", Mvar.flgEnabled, m_def_Enabled
        PropBag.WriteProperty "SavePosition", .flgSavePosition, m_def_SavePosition
        PropBag.WriteProperty "SaveSize", .flgSaveSize, m_def_SaveSize
        PropBag.WriteProperty "AutoResizeControls", .flgAutoResizeControls, m_def_AutoResizeControls
        PropBag.WriteProperty "AdditionalEvents", .flgAdditionalEvents, m_def_AdditionalEvents
        PropBag.WriteProperty "StayOnTop", .flgStayOnTop, m_def_StayOnTop
        PropBag.WriteProperty "KeepAspectRatio", .flgKeepAspectRatio, m_def_KeepAspectRatio
        PropBag.WriteProperty "FullDrag", .FullDrag, m_def_FullDrag
        PropBag.WriteProperty "CollapseButton", .flgBtnCollapse, m_def_BtnCollapse
        PropBag.WriteProperty "StayOnTopButton", .flgBtnStayOnTop, m_def_BtnStayOnTop
        PropBag.WriteProperty "FormSunken", .flgFormSunken, m_def_FormSunken
        PropBag.WriteProperty "CollapseSmallSize", .lCollapseSmallSize, m_def_CollapseSmallSize
        PropBag.WriteProperty "FormMinWidth", .lFormMinWidth, m_def_FormMinWidth
        PropBag.WriteProperty "FormMinHeight", .lFormMinHeight, m_def_FormMinHeight
        PropBag.WriteProperty "FormMinWidth", .lFormMinWidth, m_def_FormMinWidth
        PropBag.WriteProperty "FormMinHeight", .lFormMinHeight, m_def_FormMinHeight
        PropBag.WriteProperty "FormMaxWidth", .lFormMaxWidth, m_def_FormMaxWidth
        PropBag.WriteProperty "FormMaxHeight", .lFormMaxHeight, m_def_FormMaxHeight
        PropBag.WriteProperty "FormMaxPosX", .lFormMaxPosX, m_def_FormMaxPosX
        PropBag.WriteProperty "FormMaxPosY", .lFormMaxPosY, m_def_FormMaxPosY
        PropBag.WriteProperty "MsgHandle", .lMsgHandle, 0
        PropBag.WriteProperty "SaveIn", .SaveInRegOrIni, m_def_SaveIn
        PropBag.WriteProperty "BackgroundGradient", .BckgrndGradient, m_def_BackgroundGradient
        PropBag.WriteProperty "BGWidth", .BGWidth, m_def_BGWidth
        PropBag.WriteProperty "BGColorChange", .BGColorChange, m_def_BGColorChange
        PropBag.WriteProperty "BGColor1", .BGColor1, m_def_BGColor
        PropBag.WriteProperty "BGColor2", .BGColor2, m_def_BGColor
        PropBag.WriteProperty "BGColor3", .BGColor3, m_def_BGColor
        PropBag.WriteProperty "Unit", .Unit, m_def_Unit
    End With

End Sub


Private Sub UserControl_Terminate()
    ' Clean up
    
    DeactivateFunctions                         ' Stop subclassing and all other activated functions
            
End Sub



' === ALL OTHER FUNCTIONS ===

Private Sub ActivateFunctions()

    Dim oParentCtrl     As Object
    Dim lNestedLvl      As Long
    
    
    On Local Error GoTo error_handler
    
    With UserControl
        If .Ambient.UserMode = True Then
            
            ' === Looping until we get the form WizzForm is on and save a ref to it in 'frmParent'
            ' (WizzForm shouldn't be put within a container (picbox, frame,...), but who knows ... ;)
            On Error Resume Next
            Set oParentCtrl = .ParentControls.Item(1).Parent
            Do While Not (TypeOf oParentCtrl Is Form)
                Set oParentCtrl = oParentCtrl.Parent
                lNestedLvl = lNestedLvl + 1
                If oParentCtrl Is Nothing Or lNestedLvl > 10 Then     ' Something is wrong here ...
                    RaiseEvent Error(1004, "Error [Can't get reference to parent form] in procedure " + _
                            "'ActivateFunctions()' at 'ucWizzForm'")
                    
                    Exit Sub
                End If
            Loop
            Set frmParent = oParentCtrl
            If frmParent Is Nothing Then
                RaiseEvent Error(1100, "Error [Unable to activate WizzForm. No reference to parent form] in procedure " + _
                        "'ActivateFunctions()' at 'ucWizzForm'")
                
                Exit Sub
            End If
            
            On Local Error GoTo error_handler
            
    
            ' === Start subclassing for 'standard' messages
            Mvar.flgSubclassingActivated = True
            With frmParent
                Subclass_Start .hwnd

                Subclass_AddMsg .hwnd, WM_ACTIVATEAPP, MSG_BEFORE_AND_AFTER
                Subclass_AddMsg .hwnd, WM_NCACTIVATE
                Subclass_AddMsg .hwnd, WM_DISPLAYCHANGE
                Subclass_AddMsg .hwnd, WM_SYSCOLORCHANGE, MSG_BEFORE
                Subclass_AddMsg .hwnd, WM_THEMECHANGED, MSG_BEFORE

                Subclass_AddMsg .hwnd, WM_GETMINMAXINFO
                Subclass_AddMsg .hwnd, WM_ENTERSIZEMOVE
                Subclass_AddMsg .hwnd, WM_EXITSIZEMOVE
                Subclass_AddMsg .hwnd, WM_MOVING
                Subclass_AddMsg .hwnd, WM_SIZING, MSG_BEFORE
                Subclass_AddMsg .hwnd, WM_SIZE, MSG_BEFORE

                Subclass_AddMsg .hwnd, WM_MOUSEMOVE, MSG_AFTER
                Subclass_AddMsg .hwnd, WM_TIMER, MSG_BEFORE

                Subclass_AddMsg .hwnd, WM_NCPAINT
                Subclass_AddMsg .hwnd, WM_STYLECHANGED
                Subclass_AddMsg .hwnd, WM_NCLBUTTONDOWN, MSG_BEFORE
                Subclass_AddMsg .hwnd, WM_LBUTTONUP, MSG_BEFORE
                Subclass_AddMsg .hwnd, WM_NCLBUTTONDBLCLK, MSG_BEFORE
                
                Subclass_AddMsg .hwnd, WM_COPYDATA, MSG_BEFORE
            End With
            
            ' Misc others
            SetFormStayOnTop Mvar.flgStayOnTop
            
        End If
    End With
            
    
    Exit Sub


error_handler:

    RaiseEvent Error(1001, "Error [" + Err.Description + "] in procedure 'ActivateFunctions()' at 'ucWizzForm'")

End Sub

Private Function DeactivateFunctions()
        
    
    With Mvar
    
        ' Deactivate all subclassing if enabled
        If .flgSubclassingActivated = True Then
            .flgSubclassingActivated = False
            Subclass_StopAll
        End If
        
        ' Misc others
        If .flgStayOnTop = True Then
            SetFormStayOnTop False
        End If
        
    End With
        
End Function


Private Sub frmParent_Load()
    ' Here we catch the Load()event of the form

    Dim lNewVal         As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    
    
    On Local Error GoTo error_handler
    
    With frmParent
        
        ' This fixes the problem with a usercontrol on a form: Create a link to your exe, set link property e.g. to
        ' 'Show minimized' and start app using this link: Your app appears - in normal mode, NOT minimized ...
        
        Select Case StartUpWindowState()
            
            Case vbMaximizedFocus
                    .WindowState = vbMaximized
                    Mvar.flgResizeOnMaximize = True
                    
            Case vbMinimizedFocus, vbMinimizedNoFocus
                    .WindowState = vbMinimized
            
            Case vbHide
                .Visible = False
                
        End Select
        
        ' When Property set:
        SetFormSunken
        
        ' Restore forms position (Left/Top)
        If Mvar.flgSavePosition = True Then
            
            If .StartUpPosition <> vbStartUpManual Then
                RaiseEvent Error(1009, "Error [Form' s StartUpPosition not set to 'vbStartUpManual'] " + _
                        "in procedure 'frmParent_Load()' at 'ucWizzForm'")
            End If
            
            lNewVal = PersistentValueLoad(.Name + "-" + "FormPosX", -1)
            If lNewVal > -1 Then
                .Left = lNewVal
            End If
            lNewVal = PersistentValueLoad(.Name + "-" + "FormPosY", -1)
            If lNewVal > -1 Then
                .Top = lNewVal
            End If
        End If
        
        ' Restore forms size (Width/height)
        If Mvar.flgSaveSize = True Then
            If .WindowState <> vbBSNone And .WindowState <> vbSizable And .WindowState <> vbSizableToolWindow Then
                RaiseEvent Error(1010, "Error [Form' s WindowState not set to a sizeable value] " + _
                        "in procedure 'frmParent_Load()' at 'ucWizzForm'")
                
                Exit Sub
            End If
            
            lNewVal = PersistentValueLoad(.Name + "-" + "FormWidth", -1)
            If lNewVal > -1 Then
                lWidth = .Width
                .Width = lNewVal
            End If
            lNewVal = PersistentValueLoad(.Name + "-" + "FormHeight", -1)
            If lNewVal > -1 Then
                lHeight = .Height
                .Height = lNewVal
            End If
            
            ' Used to resize controls when a form is loaded
            If lWidth + lHeight Then
                ResizeControls lWidth, lHeight
            End If
            
        End If
        
    End With

    Exit Sub


error_handler:
    
    RaiseEvent Error(1002, "Error [" + Err.Description + "] in procedure 'frmParent_Load()' at 'ucWizzForm'")
    
End Sub


Private Sub frmParent_Unload(Cancel As Integer)
    ' Here we catch the Unload()event of the form
    
    On Local Error GoTo error_handler
    
    With frmParent
        
        ' === Save size/position
        If .WindowState = vbNormal Then
            
            If Mvar.flgSavePosition = True Then
                PersistentValueSave .Name + "-" + "FormPosX", .Left
                PersistentValueSave .Name + "-" + "FormPosY", .Top
            End If
            
            If Mvar.flgSaveSize = True Then
                PersistentValueSave .Name + "-" + "FormWidth", .Width
                PersistentValueSave .Name + "-" + "FormHeight", .Height
            End If
            
        End If
        
    End With


    Exit Sub


error_handler:
    
    RaiseEvent Error(1003, "Error [" + Err.Description + "] in procedure 'frmParent_QueryUnload()' at 'ucWizzForm'")

End Sub


Private Function StartUpWindowState() As VbAppWinStyle
    
    Const SW_HIDE = 0
    Const SW_MAXIMIZE = 3
    Const SW_MINIMIZE = 6
    Const SW_NORMAL = 1
    Const SW_SHOW = 5
    Const SW_SHOWDEFAULT = 10
    Const SW_SHOWMAXIMIZED = 3
    Const SW_SHOWMINIMIZED = 2
    Const SW_SHOWMINNOACTIVE = 7
    Const SW_SHOWNA = 8
    Const SW_SHOWNOACTIVATE = 4
    Const SW_SHOWNORMAL = 1
    
    Dim nInfo As tpSTARTUPINFO
    
    nInfo.cb = Len(nInfo)
    API_GetStartupInfo nInfo
    
    Select Case nInfo.wShowWindow
        
        Case SW_HIDE
                StartUpWindowState = vbHide
        
        Case SW_MAXIMIZE, SW_SHOWMAXIMIZED
                StartUpWindowState = vbMaximizedFocus
        
        Case SW_MINIMIZE, SW_SHOWMINIMIZED
                StartUpWindowState = vbMinimizedFocus
                
        Case SW_NORMAL, SW_SHOWNORMAL, SW_SHOW, SW_SHOWDEFAULT
                StartUpWindowState = vbNormalFocus
                
        Case SW_SHOWMINNOACTIVE
                StartUpWindowState = vbMinimizedNoFocus
                
        Case SW_SHOWNA, SW_SHOWNOACTIVATE
                StartUpWindowState = vbNormalNoFocus
                
    End Select
    
End Function

Private Sub SetFormSunken()

    Dim lWindowStyle    As Long
        
    With frmParent
                
        ' Get old window style
        lWindowStyle = API_GetWindowLong(.hwnd, GWL_EXSTYLE)
        
        ' Change Bit for New Style
        If Mvar.flgFormSunken = True Then
            lWindowStyle = lWindowStyle Or WS_EX_CLIENTEDGE
        Else
            lWindowStyle = lWindowStyle And Not WS_EX_CLIENTEDGE
        End If
        
        ' Set new window style
        API_SetWindowLong .hwnd, GWL_EXSTYLE, lWindowStyle

        ' Force a recalc/redraw
        .Width = .Width - 50
        .Width = .Width + 50

    End With

End Sub


Private Function PersistentValueLoad(sKey As String, Optional lDefault As Long) As Long
    ' Load values persistent saved from an INI file or from registry
    
    Const RESULT_ERROR = 0
    
    Dim sDefault        As String
    Dim sBuffer         As String
    Dim lResultLenght   As Long
    
    On Local Error GoTo error_handler
    
    If sKey = "" Then
        RaiseEvent Error(1013, "Error [Empty key value in 'PersistentValueLoad() at 'ucWizzForm'")
        
        Exit Function
    End If
    
    ' Load from INI file
    If Mvar.SaveInRegOrIni = WF_SI_INI Then
        sBuffer = Space$(20)
        sDefault = Format(lDefault) + vbNullChar
        lResultLenght = API_GetPrivateProfileString(PersistentSection, sKey, sDefault, sBuffer, Len(sBuffer), GetINIPathName())
        If lResultLenght <> RESULT_ERROR Then
            PersistentValueLoad = Val(Left$(sBuffer, lResultLenght))
        Else
            RaiseEvent Error(1008, "Error [Cannot read value (" + sKey + ") from INI file] in " + _
                    "procedure 'PersistentValueLoad()' at 'ucWizzForm'")
        End If
    
    ' Load from registry
    ElseIf Mvar.SaveInRegOrIni = WF_SI_Registry Then
        PersistentValueLoad = Val("0" + GetSetting(App.ProductName, PersistentSection, sKey, lDefault))
                
    End If

    Exit Function


error_handler:

    RaiseEvent Error(1007, "Error [" + Err.Description + "] in procedure 'PersistentValueLoad()' at 'ucWizzForm'")

End Function

Private Sub PersistentValueSave(ByVal sKey As String, ByVal lValue As Long)
    ' Save values persistent to an INI file or into registry
    
    Const RESULT_ERROR = 0
    
    Dim sValue  As String
    
    On Local Error GoTo error_handler
    
    If sKey = "" Then
        RaiseEvent Error(1013, "Error [Empty key value in 'PersistentValueSave() at 'ucWizzForm'")
        
        Exit Sub
    End If
    
    sValue = Format(lValue)
    
    ' Save into INI file
    If Mvar.SaveInRegOrIni = WF_SI_INI Then
        If API_WritePrivateProfileString(PersistentSection, sKey, sValue, GetINIPathName()) = RESULT_ERROR Then
            RaiseEvent Error(1005, "Error [Cannot write value '" + sKey + "' to INI file] in " + _
                    "procedure 'PersistentValueSave()' at 'ucWizzForm'")
        End If
    
    ' Save into registry
    ElseIf Mvar.SaveInRegOrIni = WF_SI_Registry Then
        SaveSetting App.ProductName, PersistentSection, sKey, lValue
                
    End If

    Exit Sub


error_handler:

    RaiseEvent Error(1006, "Error [" + Err.Description + "] in procedure 'PersistentValueSave()' at 'ucWizzForm'")

End Sub

Private Sub PersistentValueDelete(ByVal sKey As String)
    ' Deletes a value persistent written to an INI file or into registry
    
    Const RESULT_ERROR = 0
    
    On Local Error GoTo error_handler

    If sKey = "" Then
        RaiseEvent Error(1013, "Error [Empty key value in 'PersistentValueDelete() at 'ucWizzForm'")
        
        Exit Sub
    End If
    
    If Mvar.SaveInRegOrIni = WF_SI_INI Then
        If API_WritePrivateProfileString(PersistentSection, sKey, "", GetINIPathName()) = RESULT_ERROR Then
            RaiseEvent Error(1005, "Error [Cannot delete key '" + sKey + "' from INI file] in " + _
                    "procedure 'PersistentValueDelete()' at 'ucWizzForm'")
        End If
    
    ' Save into registry
    ElseIf Mvar.SaveInRegOrIni = WF_SI_Registry Then
        DeleteSetting App.ProductName, PersistentSection, sKey
                
    End If
    
    Exit Sub


error_handler:

    RaiseEvent Error(1012, "Error [" + Err.Description + "] in procedure 'PersistentValueDelete()' at 'ucWizzForm'")

End Sub

Private Function GetINIPathName()
    ' Change to your needs, .e.g. to Windows directory
    
    GetINIPathName = App.Path + "\" + App.EXEName + ".Ini" + vbNullChar

End Function


Private Sub ResizeControls(Optional lRestoreWidth As Long, Optional lRestoreHeight As Long)
    ' Resize special tagged controls on forms RESIZE events.
    ' This is an improved version of my earlier submission on PSC / VB.
    ' The old version has problems when the mouse is moved with fast
    ' accelleration. This version works good to me (right now ;) - let
    ' me know when you got any problems - email adress on start of this file)
    
    ' lRestoreWidth and lRestoreHeight are used on loading a form when saving of size was activated.
    ' This way not only the form keeps its size, the positions and sizes of the contained controls
    ' are handled, too.
    
    
    Static StartPosSize()   As tpAPI_RECT
    Static flgInitDone      As Boolean
    Static lOrgWidth        As Long
    Static lOrgHeight       As Long
    Static lLastWidth       As Long
    Static lLastHeight      As Long
    
    Dim lHeightChange       As Long
    Dim lHeightChangeHalf   As Long
    Dim lWidthChange        As Long
    Dim lWidthChangeHalf    As Long
    Dim lControls           As Long
    Dim i                   As Long
    Dim lPos                As Long
    Dim sTag                As String
    
    On Local Error Resume Next                              ' Avoid problems with "unusual" controls/properties ...
                                                            ' e.g. controls without width/height (like 'Timer', 'Winsock', ...)
    With frmParent
        If .WindowState = vbMinimized Then
            
            Exit Sub
        End If
        
        ' === On first start only!

        If flgInitDone = False Then
            
            ' Build an array with all (but lines) controls original positions and sizes
            lControls = .Controls.Count - 1
            ReDim StartPosSize(0 To lControls) As tpAPI_RECT
            For i = 0 To lControls
                If Not TypeOf .Controls(i) Is Line Then     ' (VB lines have X1/Y1, X2/Y2 ... :((( )
                    StartPosSize(i).lLeft = .Controls(i).Left
                    StartPosSize(i).lTop = .Controls(i).Top
                    StartPosSize(i).lRight = .Controls(i).Width
                    StartPosSize(i).lBottom = .Controls(i).Height
                End If
            Next i
                        
            ' === Set original size to build the diff
            If lRestoreWidth + lRestoreHeight Then  ' on init (when form is loaded)
                lOrgHeight = lRestoreHeight
                lOrgWidth = lRestoreWidth
            Else
                ' Save original width/height
                lOrgHeight = .Height
                lOrgWidth = .Width
            End If
            
            flgInitDone = True
            
        End If
        
        lHeightChange = .ScaleX(.Height - lOrgHeight, vbTwips, .ScaleMode)
        lWidthChange = .ScaleY(.Width - lOrgWidth, vbTwips, .ScaleMode)
        
        ' === Resize controls
        If lLastHeight <> .Height Or lLastWidth <> .Width Then    ' On change ...
            
            lWidthChangeHalf = lWidthChange / 2
            lHeightChangeHalf = lHeightChange / 2
            
            lControls = .Controls.Count - 1
            For i = 0 To lControls
                If Not TypeOf .Controls(i) Is Line Then                     ' (VB lines have X1/Y1, X2/Y2 ... :((( )
                    
                    sTag = ""                                               ' On controls without a TAG property,
                    sTag = .Controls(i).Tag                                 ' "On Error Resume Next" skips this line.

                    ' Get the part of the TAG value behind the delimiter
                    lPos = InStr(1, sTag, TAG_DELIMITER)
                    If lPos And lPos + 2 = Len(sTag) Then
                        sTag = Mid$(sTag, lPos + 1)
                        
                        If InStr(1, sTag, "L") Then
                            .Controls(i).Left = StartPosSize(i).lLeft + lWidthChange
                            
                        ElseIf InStr(1, sTag, "M") Then
                            .Controls(i).Left = StartPosSize(i).lLeft + lWidthChangeHalf
                            
                        ElseIf InStr(1, sTag, "R") Then
                            If .Controls(i).Width <> 0 Then
                                .Controls(i).Width = StartPosSize(i).lRight + lWidthChange
                            End If
                            
                        End If
                    
                        If InStr(1, sTag, "T") Then
                            .Controls(i).Top = StartPosSize(i).lTop + lHeightChange
                            
                        ElseIf InStr(1, sTag, "C") Then
                            .Controls(i).Top = StartPosSize(i).lTop + lHeightChangeHalf
                            
                        ElseIf InStr(1, sTag, "B") Then
                            If .Controls(i).Height <> 0 Then
                                .Controls(i).Height = StartPosSize(i).lBottom + lHeightChange
                            End If
                            
                        End If
                        
                    End If
                Else
                    ' Insert code for 'Line' control handling here (if you really need it ... ;) )
                End If
            Next i
        End If
        
        lLastWidth = .Width
        lLastHeight = .Height
    
    End With
    
End Sub


Private Sub SetFormStayOnTop(flgActivate As Boolean)
    ' With flgActivate = True this function ensures that the form will stay on top of all others
    ' flgActivate = False take it back to normal.
    
    If Not frmParent Is Nothing Then
        RaiseEvent FormStayOnTop(flgActivate)

        API_SetWindowPos frmParent.hwnd, ByVal IIf(flgActivate = True, HWND_TOPMOST, HWND_NOTOPMOST), _
                0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If

End Sub


Private Function SetDraggingMode(FullDrag As enFullDrag) As enFullDrag
    ' Set form to dragg with full contents or frame only
    '
    ' Parameter:    FullDrag    0 = Change nothing
    '                           1 = Drag with full contents
    '                           2 = Drag with frame only
    '
    ' Result:       Mode before call (meaning is same as parameter)
    '
    ' Notice:       Because in VB IDE we don't get WM_ACTIVATEAPP on start, this
    '               only works in compiled apps!
    '
    
    Const FULL_OFF = 0&
    Const FULL_ON = 1&
    
    Dim lResult As Long
    Dim lMode   As Long
    
    If FullDrag = WF_FD_DontChange Then
    
        Exit Function
    End If
    
    lResult = API_SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0&, lMode, 0)
    SetDraggingMode = IIf(lMode = FULL_OFF, WF_FD_No, WF_FD_Yes)
    
    If FullDrag = WF_FD_No Then
        ' Turn full dragging off
        If SetDraggingMode = True Then
            lResult = API_SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FULL_OFF, ByVal 0&, SPIF_SENDWININICHANGE)
        End If
    Else
        ' Turn full dragging on
        If SetDraggingMode = False Then
            lResult = API_SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FULL_ON, ByVal 0&, SPIF_SENDWININICHANGE)
        End If
    End If

End Function


Private Sub DrawAdditionalButtonsIntoFormCaptionbar()
    ' Main "Dispatcher"
    
    Dim RectBtn     As tpAPI_RECT
    
    With Mvar
        If .flgBtnCollapse = True Then
            GetButtonRect frmParent.hwnd, RectBtn, 4
            PaintCaptionButton .flgBtnClpsPressed, IIf(.flgCollapsed = True, BS_ArrowUp, BS_ArrowDown), RectBtn
        End If
        If .flgBtnStayOnTop = True Then
            GetButtonRect frmParent.hwnd, RectBtn, 5
            PaintCaptionButton .flgBtnSOTPressed, IIf(.flgStayOnTop = True, BS_StayOnTop, BS_DontStayOnTop), RectBtn
        End If
    End With

End Sub


Private Sub GetButtonRect(ByVal hWindow As Long, ByRef tagBtn As tpAPI_RECT, lPositionFromRight As Long)
    ' For additional buttons in caption bar:
    ' This function calculates the neede rect for the button.
    ' Written in 1/2000 by Bryan Stafford of New Vision Software - newvision@mvps.org
    ' Reformated, some changes and a fix in size (larger rect) by Light Templer
  
    Const API_FALSE     As Long = 0&
  
    Dim nHrzBdrWidth    As Long
    Dim nVrtBdrWidth    As Long
    Dim nBtnWidth       As Long
    Dim nBtnHeight      As Long
    Dim nCapHeight      As Long
    Dim nTopOffset      As Long
    Dim nPositionOffset As Long
    Dim fStyles         As Long
    Dim fExStyles       As Long
    Dim rcClient        As tpAPI_RECT
    Dim rcWindow        As tpAPI_RECT
  
    On Local Error GoTo error_handler


    ' Get the window rect and the client rect
    API_GetWindowRect hWindow, rcWindow
    API_GetClientRect hWindow, rcClient
    
    ' Get the current styles of the window
    fStyles = API_GetWindowLong(hWindow, GWL_STYLE)
    fExStyles = API_GetWindowLong(hWindow, GWL_EXSTYLE)
    
    ' Get the width/height of the respective window borders
    nHrzBdrWidth = API_GetSystemMetrics(SM_CXEDGE)
    nVrtBdrWidth = API_GetSystemMetrics(SM_CYEDGE)

    ' Determine the button and caption sizes depending on the window style and window state
    nCapHeight = API_GetSystemMetrics(SM_CYCAPTION)
    If ((fExStyles And WS_EX_TOOLWINDOW) <> 0) And (API_IsIconic(hWindow) = API_FALSE) Then
        ' Toolwindow
        nBtnWidth = API_GetSystemMetrics(SM_CXSMSIZE) - 2
        nBtnHeight = API_GetSystemMetrics(SM_CYSMSIZE) - 3
    Else
        ' Standard size
        nBtnWidth = API_GetSystemMetrics(SM_CXSIZE) - 1
        nBtnHeight = API_GetSystemMetrics(SM_CYSIZE) - 4
    End If

    ' Calculate the button size and position....
    With tagBtn
        .lLeft = (rcWindow.lRight - rcWindow.lLeft) - (((nHrzBdrWidth * 2) + 2) + (nBtnWidth * lPositionFromRight))
    
        If API_IsIconic(hWindow) Then
            nTopOffset = (((rcWindow.lBottom - rcWindow.lTop) \ 2) - (nBtnHeight \ 2)) - (nVrtBdrWidth + 2)
            nPositionOffset = 1
        Else
            nTopOffset = ((((rcWindow.lBottom - rcWindow.lTop) - rcClient.lBottom) - (nVrtBdrWidth * 2)) - nCapHeight) \ 2
            ' !!! When 'FormSunken' is set 'rcClient.lBottom' doesn't go below zero (maybe a Windows bug) so
            '     additionl buttons are jumping higher. Set min windows size to avaoid this when you need sunken forms.
            nPositionOffset = IIf(fStyles And WS_THICKFRAME, 2, 1)
        End If
    
        If lPositionFromRight > 1 Then
            .lLeft = .lLeft - nPositionOffset
        End If
        
        .lLeft = .lLeft + 1                         ' Added because imho needed ...
        .lTop = (nVrtBdrWidth + 2) + nTopOffset
        .lRight = .lLeft + nBtnWidth
        .lBottom = .lTop + nBtnHeight
        
        If ((fExStyles And WS_EX_TOOLWINDOW) <> 0) Then
            ' Adjustments to toolwindows
            .lTop = .lTop + 2
            .lBottom = .lBottom + 1
            .lLeft = .lLeft + 1
        End If
    
        If ((fExStyles And WS_EX_CLIENTEDGE) <> 0) Then
            ' Adjustments to windows with 'sunken' style
            .lTop = .lTop - 2
            .lBottom = .lBottom - 2
        End If
        
    End With

    Exit Sub


error_handler:
    
    RaiseEvent Error(1011, "Error [" + Err.Description + "] in procedure 'GetButtonRect()' at 'ucWizzForm'")
      
End Sub


Private Sub GetScreenPoint(ByVal hWindow As Long, ByVal lParam As Long, ByRef pt As tpAPI_POINT)
    ' Converts a local point from a hittest to a screen point
    
    With pt
        .X = LoWord(lParam)
        .Y = HiWord(lParam)
    End With
    API_ClientToScreen hWindow, pt

End Sub


Private Sub PaintCaptionButton(flgBtnPressed As Boolean, BtnSymbol As enBtnSymbol, ByRef RectButton As tpAPI_RECT)
    ' Here the drawing of the symbol into the button is done by using Windows' 'Marlett' font (a symbol font)
    ' For Windows XP this needs some modifications! - I'm using NT4, W2K and Win98 and don't have access
    ' to check it on Win XP. Any improvements are welcome - please feel free to send me an email - adress on start of this code.
        
    Const API_FALSE As Long = 0&
    Const TRANSPARENT = 1&

    Dim hNonClientDC    As Long
    Dim lState          As Long
    Dim hOldFont        As Long
    Dim hNewFont        As Long
    Dim fExStyles       As Long
    Dim sSign           As String
    
    
    With frmParent
        If API_IsIconic(.hwnd) = API_FALSE Then
            ' Draw onto forms caption
            hNonClientDC = API_GetWindowDC(.hwnd)

            ' Button frame
            lState = IIf(flgBtnPressed = True, DFCS_BUTTONPUSH Or DFCS_PUSHED, DFCS_BUTTONPUSH)
            API_DrawFrameControl hNonClientDC, RectButton, DFC_BUTTON, lState

            ' Button contents
            hNewFont = BuildMarlettFont(RectButton.lBottom - (RectButton.lTop + 1))
            hOldFont = API_SelectObject(hNonClientDC, hNewFont)
            fExStyles = API_GetWindowLong(frmParent.hwnd, GWL_EXSTYLE)
            If ((fExStyles And WS_EX_TOOLWINDOW) <> 0) And (API_IsIconic(frmParent.hwnd) = API_FALSE) Then
                ' Toolwindow
                ' API_OffsetRect RectButton, 0, 0       ' Nothing to do in NT 4 ..., but prepared for changes ...
            Else
                API_OffsetRect RectButton, 1, 1
            End If
            API_SetBkMode hNonClientDC, TRANSPARENT
            sSign = Choose(BtnSymbol, "5", "6", "i", "n")
            API_DrawText hNonClientDC, sSign, 1, RectButton, 0&
            API_SelectObject hNonClientDC, hOldFont
            API_DeleteObject hNewFont
        End If
    End With
  
End Sub


Private Function BuildMarlettFont(lFontSize As Long) As Long
    ' Get handle to this special window symbols font
    
    Const SYMBOL_CHARSET = 2
    
    Dim TheFont As tpLOGFONT
    
    With TheFont
        .lfFaceName = "Marlett" + vbNullChar
        .lfHeight = lFontSize
        .lfCharSet = SYMBOL_CHARSET     ' Important ... !
    End With
    BuildMarlettFont = API_CreateFontIndirect(TheFont)

End Function


Private Sub CollapseForm()
    ' Here we shrink or expand forms height.
    ' Keep 'Min Size' in mind ...
    
    Dim lNewSize As Long
    
    RaiseEvent FormCollapse(Mvar.flgBtnCollapse)
    
    With Mvar
        If .flgCollapsed = True Then
            .lCollapseOrgSize = frmParent.Height
            
            lNewSize = .lCollapseSmallSize
            If .Unit = WF_UN_Pixels Then
                lNewSize = frmParent.ScaleY(lNewSize, vbPixels, vbTwips)
            End If
            
            If .flgFormSunken = True And lNewSize < 1 Then
                lNewSize = 30
            End If
            
            frmParent.Height = lNewSize
        Else
            frmParent.Height = .lCollapseOrgSize
        End If
    End With
    
End Sub


Private Sub DrawBackgroundGradient(Optional flgRedrawFullForm As Boolean = False)
    ' Draw a gradient on persistent bitmap of form.
    '
    ' Important:    Set Form.AutoRedraw = True !
        
    ' Notice:       Why I havn't implemented Left/Right gradients here, too? Well, ;) ... imho ... most of times they
    '               result in an ugly looking screen design. But of course, if you really need them it should be easy
    '               for you to add them here.
        
    Dim FormRect    As tpAPI_RECT
    Dim lColor1     As Long
    Dim lColor2     As Long
    Dim lColor3     As Long
    

    On Local Error GoTo error_handler
    
        
    ' Get drawing rectangle (width/height) in pixels the API way ;)
    API_GetClientRect frmParent.hwnd, FormRect
    With FormRect
        .lRight = .lRight - .lLeft
        .lLeft = 0
        .lBottom = .lBottom - .lTop
        .lTop = 0
    End With
    
    With Mvar
        If .BGWidth > 0 Then
            ' Not the whole form from left to right, we stop after 'BGWidth' pixels with drawing
            FormRect.lRight = .BGWidth
        End If
            
        If flgRedrawFullForm = True Then
            frmParent.Picture = Nothing
            frmParent.Cls
        End If
            
            
        ' Prepare API colors
        lColor1 = OLEColorToRGB(.BGColor1)
        lColor2 = OLEColorToRGB(.BGColor2)
        lColor3 = OLEColorToRGB(.BGColor3)
        
        ' Design gradients
        Select Case .BckgrndGradient
    
            Case enWFGradient.WF_GR_None
                    
                    Exit Sub
    
            
            ' === Here only the first and the 2nd color are used to draw the 'standard' background gradient.
            Case enWFGradient.WF_GR_TwoColorGradient, enWFGradient.WF_GR_TwoColorGradPlusBlock
                    
                    ' For a 'TwoColorGradient' we fill the bottom area with the forms backcolor
                    If .BckgrndGradient = enWFGradient.WF_GR_TwoColorGradient Then
                        lColor3 = OLEColorToRGB(frmParent.BackColor)
                    End If
                    
                    If .BGColorChange > 0 And .BGColorChange < 100 Then
                        
                        ' = We have a PERCENT value. Draw gradient from top downto this value. (e.g. 50% means: to middle of form)
                        
                        ' Get border
                        FormRect.lTop = FormRect.lBottom * (.BGColorChange / 100)
                                               
                        ' Draw bottom area (always neccessary, resizing ...)
                        DrawRect frmParent.hdc, FormRect, lColor3
                        
                        ' Adjust for gradient in top area
                        FormRect.lBottom = FormRect.lTop
                        FormRect.lTop = 0
                        
                    ElseIf .BGColorChange < 0 Then
                        
                        ' = We have a NUMBER for lines of pixel. Draw gradient from top with this number of lines. Rest is empty.
                        ' Get border
                        FormRect.lTop = -1 * .BGColorChange + 1
                                               
                        ' Refill bottom area (resizing ...)
                        DrawRect frmParent.hdc, FormRect, lColor3
                        
                        ' Adjust for gradient in top area
                        FormRect.lTop = 0
                        FormRect.lBottom = -1 * .BGColorChange
                        DrawTopDownGradient frmParent.hdc, FormRect, lColor1, lColor2
                        
                    End If
                        
                    DrawTopDownGradient frmParent.hdc, FormRect, lColor1, lColor2
                    
                    
                
            ' === Here all three colors are used to draw two background gradients.
            Case enWFGradient.WF_GR_ThreeColorGradient
            
                    
                    If .BGColorChange > 0 And .BGColorChange <= 100 Then
                        
                        ' = We have a PERCENT value. Draw first gradient from top downto this value.
                        '   (e.g. 50% means: to middle of form)
                        
                        ' Get border
                        FormRect.lTop = FormRect.lBottom * (.BGColorChange / 100)

                        ' Draw a gradient into bottom area
                        DrawTopDownGradient frmParent.hdc, FormRect, lColor2, lColor3
                                                
                        ' Adjust for top area
                        FormRect.lBottom = FormRect.lTop + 1
                        FormRect.lTop = 0
                        
                        ' Draw a gradient into top area
                        DrawTopDownGradient frmParent.hdc, FormRect, lColor1, lColor2
                       
                    ElseIf .BGColorChange < 0 Then
                        
                        ' = We have a NUMBER for lines of pixel. Draw gradient from top with
                        '   this number of lines. Below is the 2nd gradient.
                        
                        ' Get border
                        FormRect.lTop = -1 * .BGColorChange
                                               
                        ' Draw a gradient into bottom area
                        DrawTopDownGradient frmParent.hdc, FormRect, lColor2, lColor3
                        
                        ' Adjust for gradient in top area
                        FormRect.lTop = 0
                        FormRect.lBottom = -1 * .BGColorChange
                        
                        ' Draw a gradient into top area
                        DrawTopDownGradient frmParent.hdc, FormRect, lColor1, lColor2
                        
                    End If
                        
        End Select
        
        ' Copy to make gradient persistent
        frmParent.Picture = frmParent.Image
        
    End With
    
    
    Exit Sub


error_handler:
        
End Sub

Private Function OLEColorToRGB(ByVal oColor As OLE_COLOR) As Long
    ' Convert color values from OLE representation to RGB representation
    
    Dim lRGB    As Long
    Dim hPal    As Long
    
    OLEColorToRGB = IIf(API_OleTranslateColor(oColor, hPal, lRGB), API_INVALID_COLOR, lRGB)
    
End Function

Private Sub DrawRect(hdc As Long, TheRect As tpAPI_RECT, lRGBFillColor As Long)
    ' Draw a rectangle
    
    Dim lBrush  As Long
    
    If lRGBFillColor <> API_INVALID_COLOR Then
        lBrush = API_CreateSolidBrush(lRGBFillColor)
        API_FillRect hdc, TheRect, lBrush
        API_DeleteObject lBrush
    End If
    
End Sub


Private Sub DrawTopDownGradient(hdc As Long, rc As tpAPI_RECT, ByVal lRGBColorFrom As Long, ByVal lRGBColorTo As Long)
    ' This sub is the result from the small competition I started on PSC VB to find the fastest compatible
    ' gradient sub. Thx alot to Carles P.V. for this jewel!
    ' Plz look at  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57192&lngWId=1  for details.
    
    Dim uBIH            As tpBITMAPINFOHEADER
    Dim lBits()         As Long
    Dim lColor          As Long
    
    Dim X               As Long
    Dim Y               As Long
    Dim xEnd            As Long
    Dim yEnd            As Long
    Dim ScanlineWidth   As Long
    Dim yOffset         As Long
    
    Dim R               As Long
    Dim G               As Long
    Dim B               As Long
    Dim end_R           As Long
    Dim end_G           As Long
    Dim end_B           As Long
    Dim dR              As Long
    Dim dG              As Long
    Dim dB              As Long
    
    
    ' Split a RGB long value into components - FROM gradient color
    lRGBColorFrom = lRGBColorFrom And &HFFFFFF                      ' "SplitRGB"  by www.Abstractvb.com
    R = lRGBColorFrom Mod &H100&                                    ' Should be the fastest way in pur VB
    lRGBColorFrom = lRGBColorFrom \ &H100&                          ' See test on VBSpeed (http://www.xbeat.net/vbspeed/)
    G = lRGBColorFrom Mod &H100&                                    ' Btw: API solution with RTLMoveMem is slower ... ;)
    lRGBColorFrom = lRGBColorFrom \ &H100&
    B = lRGBColorFrom Mod &H100&
    
    ' Split a RGB long value into components - TO gradient color
    lRGBColorTo = lRGBColorTo And &HFFFFFF
    end_R = lRGBColorTo Mod &H100&
    lRGBColorTo = lRGBColorTo \ &H100&
    end_G = lRGBColorTo Mod &H100&
    lRGBColorTo = lRGBColorTo \ &H100&
    end_B = lRGBColorTo Mod &H100&
    
    
    '-- Loops bounds
    xEnd = rc.lRight - rc.lLeft
    yEnd = rc.lBottom - rc.lTop
    
    ' Check:  Top lower than Bottom ?
    If yEnd < 1 Then
    
        Exit Sub
    End If
    
    '-- Scanline width
    ScanlineWidth = xEnd + 1
    yOffset = -ScanlineWidth
    
    '-- Initialize array size
    ReDim lBits((xEnd + 1) * (yEnd + 1) - 1) As Long
       
    '-- Get color distances
    dR = end_R - R
    dG = end_G - G
    dB = end_B - B
       
    '-- Gradient loop over rectangle
    For Y = 0 To yEnd
        
        '-- Calculate color and *y* offset
        lColor = B + (dB * Y) \ yEnd + 256 * (G + (dG * Y) \ yEnd) + 65536 * (R + (dR * Y) \ yEnd)
        
        yOffset = yOffset + ScanlineWidth
        
        '-- *Fill* line
        For X = yOffset To xEnd + yOffset
            lBits(X) = lColor
        Next X
        
    Next Y
    
    '-- Prepare bitmap info structure
    With uBIH
        .biSize = Len(uBIH)
        .biBitCount = 32
        .biPlanes = 1
        .biWidth = xEnd + 1
        .biHeight = -yEnd + 1
    End With
    
    '-- Finaly, paint *bits* onto given DC
    API_StretchDIBits hdc, _
            rc.lLeft, rc.lTop, _
            xEnd, yEnd, _
            0, 0, _
            xEnd, yEnd, _
            lBits(0), _
            uBIH, _
            API_DIB_RGB_COLORS, _
            vbSrcCopy

End Sub


Private Function ValToUsedUnit(lValue As Long)
    ' Convert an internal value to the unit the user of WizzForm has choosen
    ' We can work (right now) with 'Pixels' and 'Twips'
    ' Add more code here if you need e.g. Inch or something else
    ' Notice: Internaly all values are saved in 'Pixels' !
    
    With Mvar
        If .Unit = WF_UN_Pixels Then
            ValToUsedUnit = lValue
            
        ElseIf .Unit = WF_UN_Twips Then
            ValToUsedUnit = UserControl.ScaleX(lValue, vbPixels, vbTwips)       ' I never saw a difference for ScaleX and ScaleY
                                                                                ' in real world situations, so this is ok for me ;)
        End If
    End With
    
End Function

Private Function UsedUnitToValue(lValue As Long)
    ' Convert a value in the unit the user of WizzForm has choosen to pixels
    ' We can work (right now) with 'Pixels' and 'Twips'
    ' Add more code here if you need e.g. Inch or something else
    ' Notice: Internaly all values are saved in 'Pixels' !
    
    With Mvar
        If .Unit = WF_UN_Pixels Then
            UsedUnitToValue = lValue
            
        ElseIf .Unit = WF_UN_Twips Then
            UsedUnitToValue = UserControl.ScaleX(lValue, vbTwips, vbPixels)     ' I never saw a difference for ScaleX and ScaleY
                                                                                ' in real world situations, so this is ok for me ;)
        End If
    End With
    
End Function





' ======================================================================================
' = Subclass code - The programmer may call any of the following Subclass_??? routines =
' ======================================================================================


Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As enMsgWhen = MSG_AFTER)
    ' Add a message to the list of those that will invoke a callback.
    ' You should Subclass_Start first and then add the messages
    '
    ' Parameters:
    '       lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
    '       uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
    '       When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    
    With sc_aSubData(zIdx(lng_hWnd))
        If When And enMsgWhen.MSG_BEFORE Then
            zAddMsg uMsg, .aMsgTblB, .nMsgCntB, enMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And enMsgWhen.MSG_AFTER Then
            zAddMsg uMsg, .aMsgTblA, .nMsgCntA, enMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With
    
End Sub


Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As enMsgWhen = MSG_AFTER)
    ' Delete a message from the table of those that will invoke a callback.
    ' Parameters:
    '       lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    '       uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    '       When      - Whether the msg is to be removed from the before, after or both callback tables
    
    With sc_aSubData(zIdx(lng_hWnd))
        If When And enMsgWhen.MSG_BEFORE Then
            zDelMsg uMsg, .aMsgTblB, .nMsgCntB, enMsgWhen.MSG_BEFORE, .nAddrSub
        End If
        If When And enMsgWhen.MSG_AFTER Then
            zDelMsg uMsg, .aMsgTblA, .nMsgCntA, enMsgWhen.MSG_AFTER, .nAddrSub
        End If
    End With

End Sub


Private Function Subclass_InIDE() As Boolean
    ' Return whether we're running in the IDE.
    
    Debug.Assert zSetTrue(Subclass_InIDE)
    
End Function


Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    ' Start subclassing the passed window handle
    '
    ' Parameters:
    '               lng_hWnd  - The handle of the window to be subclassed
    ' Returns:
    '               The sc_aSubData() index
    
    
    Const CODE_LEN              As Long = 204                             ' Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"             ' We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                      ' VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"              ' SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                      ' Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                        ' Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                        ' Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                              ' Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                              ' Address of the previous WndProc
    Const PATCH_03              As Long = 78                              ' Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                             ' Address of the previous WndProc
    Const PATCH_07              As Long = 121                             ' Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                             ' Address of the owner object
    
    Static aBuf(1 To CODE_LEN)  As Byte                                   ' Static code buffer byte array
    Static pCWP                 As Long                                   ' Address of the CallWindowsProc
    Static pEbMode              As Long                                   ' Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                   ' Address of the SetWindowsLong function
    
    Dim i                       As Long                                   ' Loop index
    Dim J                       As Long                                   ' Loop index
    Dim nSubIdx                 As Long                                   ' Subclass data index
    Dim sHex                    As String                                 ' Hex code string
  
    ' If it's the first time through here..
    If aBuf(1) = 0 Then
  
        ' The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
                "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
                "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
                "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

        ' Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While J < CODE_LEN
            J = J + 1
            aBuf(J) = Val("&H" & Mid$(sHex, i, 2))                          ' Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                ' Next pair of hex characters
    
        ' Get API function addresses
        If Subclass_InIDE Then                                              ' If we're running in the VB IDE
            aBuf(16) = &H90                                                 ' Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                 ' Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                         ' Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                             ' Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                     ' VB5 perhaps
            End If
        End If
    
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                ' Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                ' Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                               ' Create the first sc_aSubData element
        
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                ' If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                             ' Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData            ' Create a new sc_aSubData element
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lng_hWnd                                                    ' Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                       ' Allocate memory for the machine code WndProc
        .nAddrOrig = API_SetWindowLong(.hwnd, GWL_WNDPROC, .nAddrSub)       ' Set our WndProc in place
        API_RtlMoveMemory ByVal .nAddrSub, aBuf(1), CODE_LEN                ' Copy the machine code from the static byte array to the code array in sc_aSubData
        zPatchRel .nAddrSub, PATCH_01, pEbMode                              ' Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        zPatchVal .nAddrSub, PATCH_02, .nAddrOrig                           ' Original WndProc address for CallWindowProc, call the original WndProc
        zPatchRel .nAddrSub, PATCH_03, pSWL                                 ' Patch the relative address of the SetWindowLongA api function
        zPatchVal .nAddrSub, PATCH_06, .nAddrOrig                           ' Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        zPatchRel .nAddrSub, PATCH_07, pCWP                                 ' Patch the relative address of the CallWindowProc api function
        zPatchVal .nAddrSub, PATCH_0A, ObjPtr(Me)                           ' Patch the address of this object instance into the static machine code buffer
    End With
  
End Function


Private Sub Subclass_StopAll()
    ' Stop all subclassing
  
    Dim i As Long
    
    i = UBound(sc_aSubData())                                               ' Get the upper bound of the subclass data array
    Do While i >= 0                                                         ' Iterate through each element
        With sc_aSubData(i)
            If .hwnd <> 0 Then                                              ' If not previously Subclass_Stop'd
                Subclass_Stop .hwnd                                         ' Subclass_Stop
            End If
        End With
    
        i = i - 1                                                           ' Next element
    Loop
    
End Sub


Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    ' Stop subclassing the passed window handle
    '
    ' Parameters:
    '       lng_hWnd  - The handle of the window to stop being subclassed
    
    
    With sc_aSubData(zIdx(lng_hWnd))
        API_SetWindowLong .hwnd, GWL_WNDPROC, .nAddrOrig                    ' Restore the original WndProc
        zPatchVal .nAddrSub, PATCH_05, 0                                    ' Patch the Table B entry count to ensure no further 'before' callbacks
        zPatchVal .nAddrSub, PATCH_09, 0                                    ' Patch the Table A entry count to ensure no further 'after' callbacks
        GlobalFree .nAddrSub                                                ' Release the machine code memory
        .hwnd = 0                                                           ' Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                       ' Clear the before table
        .nMsgCntA = 0                                                       ' Clear the after table
        Erase .aMsgTblB                                                     ' Erase the before table
        Erase .aMsgTblA                                                     ' Erase the after table
    End With

End Sub



' ============================================================================
' = These z??? routines are exclusively called by the Subclass_??? routines. =
' ============================================================================

Private Sub zAddMsg(ByVal uMsg As Long, _
                    ByRef aMsgTbl() As Long, _
                    ByRef nMsgCnt As Long, _
                    ByVal When As enMsgWhen, _
                    ByVal nAddr As Long)
                    
    ' Worker sub for Subclass_AddMsg
  
    Dim nEntry  As Long                                                     ' Message table entry index
    Dim nOff1   As Long                                                     ' Machine code buffer offset 1
    Dim nOff2   As Long                                                     ' Machine code buffer offset 2

    If uMsg = ALL_MESSAGES Then                                             ' If all messages
        nMsgCnt = ALL_MESSAGES                                              ' Indicates that all messages will callback
    Else                                                                    ' Else a specific message number
        Do While nEntry < nMsgCnt                                           ' For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
    
            If aMsgTbl(nEntry) = 0 Then                                     ' This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                      ' Re-use this entry
                
                Exit Sub                                                    ' Bail
    
            ElseIf aMsgTbl(nEntry) = uMsg Then                              ' The msg is already in the table!
                
                Exit Sub                                                    ' Bail
            End If
        Loop                                                                ' Next entry
    
        nMsgCnt = nMsgCnt + 1                                               ' New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                        ' Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                             ' Store the message number in the table
    End If
    
    If When = enMsgWhen.MSG_BEFORE Then                                     ' If before
        nOff1 = PATCH_04                                                    ' Offset to the Before table
        nOff2 = PATCH_05                                                    ' Offset to the Before table entry count
    Else                                                                    ' Else after
        nOff1 = PATCH_08                                                    ' Offset to the After table
        nOff2 = PATCH_09                                                    ' Offset to the After table entry count
    End If
    
    If uMsg <> ALL_MESSAGES Then
        zPatchVal nAddr, nOff1, VarPtr(aMsgTbl(1))                          ' Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    zPatchVal nAddr, nOff2, nMsgCnt                                         ' Patch the appropriate table entry count
  
End Sub


Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    ' Return the memory address of the passed function in the passed dll
    
    zAddrFunc = API_GetProcAddress(API_GetModuleHandle(sDLL), sProc)
    Debug.Assert zAddrFunc                                                  ' You may wish to comment out this line
                                                                            ' if you're using vb5 else the EbMode
                                                                            ' GetProcAddress will stop here everytime
                                                                            ' because we look for vba6.dll first
End Function


Private Sub zDelMsg(ByVal uMsg As Long, _
                    ByRef aMsgTbl() As Long, _
                    ByRef nMsgCnt As Long, _
                    ByVal When As enMsgWhen, _
                    ByVal nAddr As Long)
                    
    ' Worker sub for Subclass_DelMsg
    
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                                             ' If deleting all messages
        nMsgCnt = 0                                                         ' Message count is now zero
        If When = enMsgWhen.MSG_BEFORE Then                                 ' If before
            nEntry = PATCH_05                                               ' Patch the before table message count location
        Else                                                                ' Else after
            nEntry = PATCH_09                                               ' Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                    ' Patch the table message count to zero
    Else                                                                    ' Else deleteting a specific message
        Do While nEntry < nMsgCnt                                           ' For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                  ' If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                         ' Mark the table slot as available
                
                Exit Do                                                     ' Bail
            End If
        Loop                                                                ' Next entry
    End If
  
End Sub


Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    ' Get the sc_aSubData() array index of the passed hWnd
    ' Get the upper bound of sc_aSubData() - If you get an error here, you're probably
    ' Subclass_AddMsg-ing before Subclass_Start
    
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                      ' Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hwnd = lng_hWnd Then                                        ' If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                            ' If we're searching not adding
                    
                    Exit Function                                           ' Found
                End If
            
            ElseIf .hwnd = 0 Then                                           ' If this an element marked for reuse.
                If bAdd Then                                                ' If we're adding
                    
                    Exit Function                                           ' Re-use it
                End If
                
            End If
        End With
        zIdx = zIdx - 1                                                     ' Decrement the index
    Loop
    
    If Not bAdd Then
        Debug.Assert False                                                  ' hWnd not found, programmer error
    End If

    ' If we exit here, we're returning -1, no freed elements were found

End Function


Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    ' Patch the machine code buffer at the indicated offset with the relative address to the target address.
  
    API_RtlMoveMemory ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4
    
End Sub


Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    ' Patch the machine code buffer at the indicated offset with the passed value
    
    API_RtlMoveMemory ByVal nAddr + nOffset, nValue, 4
    
End Sub


Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    ' Worker function for Subclass_InIDE
    
    zSetTrue = True
    bValue = True
    
End Function

' ********************************************************************
' *                       END OF SUBCLASS CODE                       *
' ********************************************************************





' *************************************
' *           PROPERTIES              *
' *************************************

Public Property Get SavePosition() As Boolean
Attribute SavePosition.VB_Description = "Save/restore last position of form."
    ' Save forms position persistent
    
    SavePosition = Mvar.flgSavePosition

End Property

Public Property Let SavePosition(ByVal flgNewSavePosition As Boolean)
    ' Save forms position persistent
    
    Mvar.flgSavePosition = flgNewSavePosition
    PropertyChanged "SavePosition"
    
End Property


Public Property Get SaveSize() As Boolean
Attribute SaveSize.VB_Description = "Save/restore last size of form."
    ' Save forms size persistent
    
    SaveSize = Mvar.flgSaveSize

End Property

Public Property Let SaveSize(ByVal flgNewSaveSize As Boolean)
    ' Save forms size persistent
    
    Mvar.flgSaveSize = flgNewSaveSize
    PropertyChanged "SaveSize"
    
End Property


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Switch on/off all functions of WizzForm"
Attribute Enabled.VB_UserMemId = -514
    ' Enable/disable ALL functions of WizzForm
        
    Enabled = Mvar.flgEnabled

End Property

Public Property Let Enabled(ByVal flgNew_Enabled As Boolean)
    ' Enable/disable ALL functions of WizzForm
    
    With Mvar
        ' When value really changes only
        If flgNew_Enabled = True And .flgEnabled = False Then
            ActivateFunctions
            
        ElseIf flgNew_Enabled = False And .flgEnabled = True Then
            DeactivateFunctions
            
        End If
            
        .flgEnabled = flgNew_Enabled
        PropertyChanged "Enabled"
    End With
        
End Property


Public Property Get FormMinWidth() As Long
Attribute FormMinWidth.VB_Description = "Min width of form"
    ' Limit forms width to a min value
    
    FormMinWidth = ValToUsedUnit(Mvar.lFormMinWidth)
            
End Property

Public Property Let FormMinWidth(ByVal lNew_FormMinWidth As Long)
    ' Limit forms width to a min value
    
    If lNew_FormMinWidth >= 0 Then
        Mvar.lFormMinWidth = UsedUnitToValue(lNew_FormMinWidth)
        PropertyChanged "FormMinWidth"
    End If
    
End Property


Public Property Get FormMinHeight() As Long
Attribute FormMinHeight.VB_Description = "Min high of form"
    ' Limit forms height to a min value
    
    FormMinHeight = ValToUsedUnit(Mvar.lFormMinHeight)
    
End Property

Public Property Let FormMinHeight(ByVal lNew_FormMinHeight As Long)
    ' Limit forms height to a min value
    
    If lNew_FormMinHeight >= 0 Then
        Mvar.lFormMinHeight = UsedUnitToValue(lNew_FormMinHeight)
        PropertyChanged "FormMinHeight"
    End If
    
End Property


Public Property Get FormMaxWidth() As Long
Attribute FormMaxWidth.VB_Description = "Max width of form"
    ' Limit forms height to a max value
    
    FormMaxWidth = ValToUsedUnit(Mvar.lFormMaxWidth)
    
End Property

Public Property Let FormMaxWidth(ByVal lNew_FormMaxWidth As Long)
    ' Limit forms height to a max value
    
    If lNew_FormMaxWidth >= 0 And lNew_FormMaxWidth >= Mvar.lFormMinWidth Then
        Mvar.lFormMaxWidth = UsedUnitToValue(lNew_FormMaxWidth)
        PropertyChanged "FormMaxWidth"
    End If
    
End Property


Public Property Get FormMaxHeight() As Long
Attribute FormMaxHeight.VB_Description = "Max high of form"
    ' Limit forms height to a max value
    
    FormMaxHeight = ValToUsedUnit(Mvar.lFormMaxHeight)
    
End Property

Public Property Let FormMaxHeight(ByVal lNew_FormMaxHeight As Long)
    ' Limit forms height to a max value
    
    If lNew_FormMaxHeight >= 0 And lNew_FormMaxHeight >= Mvar.lFormMinHeight Then
        Mvar.lFormMaxHeight = UsedUnitToValue(lNew_FormMaxHeight)
        PropertyChanged "FormMaxHeight"
    End If
    
End Property


Public Property Get FormMaxPosX() As Long
Attribute FormMaxPosX.VB_Description = "When form's max width is smaller than screen area and you klick the maximize button, form is moved to this X position."
    ' X position of form when 'maximize form' is selected and forms max size is set to a value smaller than screens size
    
    FormMaxPosX = ValToUsedUnit(Mvar.lFormMaxPosX)
    
End Property

Public Property Let FormMaxPosX(ByVal lNew_FormMaxPosX As Long)
    ' X position of form when 'maximize form' is selected and forms max size is set to a value smaller than screens size
    
    If lNew_FormMaxPosX >= 0 Then
        Mvar.lFormMaxPosX = UsedUnitToValue(lNew_FormMaxPosX)
        PropertyChanged "FormMaxPosX"
    End If
    
End Property


Public Property Get FormMaxPosY() As Long
Attribute FormMaxPosY.VB_Description = "When form's max width is smaller than screen area and you klick the maximize button, form is moved to this Y position."
    ' Y position of form when 'maximize form' is selected and forms max size is set to a value smaller than screens size
    
    FormMaxPosY = ValToUsedUnit(Mvar.lFormMaxPosY)
    
End Property

Public Property Let FormMaxPosY(ByVal lNew_FormMaxPosY As Long)
    ' Y position of form when 'maximize form' is selected and forms max size is set to a value smaller than screens size
    
    If lNew_FormMaxPosY >= 0 Then
        Mvar.lFormMaxPosY = UsedUnitToValue(lNew_FormMaxPosY)
        PropertyChanged "FormMaxPosY"
    End If
    
End Property


Public Property Get AutoResizeControls() As Boolean
Attribute AutoResizeControls.VB_Description = "Controls with special values in their tag property will automaticly resized when form size changes."
    ' Resize tagged controls
    
    AutoResizeControls = Mvar.flgAutoResizeControls
    
End Property

Public Property Let AutoResizeControls(ByVal flgNew_AutoResizeControls As Boolean)
    ' Resize tagged controls
    
    Mvar.flgAutoResizeControls = flgNew_AutoResizeControls
    PropertyChanged "AutoResizeControls"
    
End Property


Public Property Get AdditionalEvents() As Boolean
Attribute AdditionalEvents.VB_Description = "Raise events you don't have on pure forms (App activated, form moved/sized, ...)"
    ' Switch raising of addtional form events on/off
    
    AdditionalEvents = Mvar.flgAdditionalEvents
    
End Property

Public Property Let AdditionalEvents(ByVal flgNew_AdditionalEvents As Boolean)
    ' Switch raising of addtional form events on/off
    
    Mvar.flgAdditionalEvents = flgNew_AdditionalEvents
    PropertyChanged "AdditionalEvents"
    
End Property


Public Property Get SaveIn() As enSaveIn
Attribute SaveIn.VB_Description = "Select storage type for forms size and position: An INI file or the registry."
    ' Saving forms position/size to registry or into an Ini file
    
    SaveIn = Mvar.SaveInRegOrIni
    
End Property

Public Property Let SaveIn(ByVal enNew_SaveIn As enSaveIn)
    ' Saving forms position/size to registry or into an Ini file
    
    Mvar.SaveInRegOrIni = enNew_SaveIn
    PropertyChanged "SaveIn"
    
End Property


Public Property Get StayOnTop() As Boolean
Attribute StayOnTop.VB_Description = "Let form sweep over all other forms even when not selected."
    ' This "classicer" shouldn't miss ;) - The form stays on top of all other windows
    
    StayOnTop = Mvar.flgStayOnTop
    
End Property

Public Property Let StayOnTop(ByVal flgNew_StayOnTop As Boolean)
    ' This "classicer" cannot miss ;) - The form keeps staying on top of all other windows
    
    Mvar.flgStayOnTop = flgNew_StayOnTop
    PropertyChanged "StayOnTop"
    
    SetFormStayOnTop Mvar.flgStayOnTop
    
End Property


Public Property Get KeepAspectRatio() As Boolean
Attribute KeepAspectRatio.VB_Description = "Switched on resizing the form keeps ratio height/width."
    ' Keep aspect ratio on form resizing
    
    KeepAspectRatio = Mvar.flgKeepAspectRatio
    
End Property

Public Property Let KeepAspectRatio(ByVal flgNew_KeepAspectRatio As Boolean)
    ' Keep aspect ratio on form resizing
    
    Mvar.flgKeepAspectRatio = flgNew_KeepAspectRatio
    PropertyChanged "KeepAspectRatio"
    
End Property


Public Property Get FullDrag() As enFullDrag
Attribute FullDrag.VB_Description = "Drag form with contents or with frame only."
    ' Drag form with full contents or with frame only - ignoring system wide setting.
    '
    ' Hint:         Because in VB IDE we don't get WM_ACTIVATEAPP on start, this
    '               only works in compiled apps!
    '
    ' Attention:    Only set this value on your first (main) form!
    '               Leave this property on all other forms on default (DontChange = 0) !
    
    FullDrag = Mvar.FullDrag
    
End Property

Public Property Let FullDrag(ByVal New_FullDrag As enFullDrag)
    ' Drag form with full contents or with frame only - ignoring system wide setting.
    '
    ' Notice:       Because of in VB IDE we don't get WM_ACTIVATEAPP on start, this
    '               only works in compiled apps!
    '
    ' Attention:  Only set this value on your first (main) form!
    '             Leave this property on all other forms on default (DontChange = 0) !
    
    
    Mvar.FullDrag = New_FullDrag
    PropertyChanged "FullDrag"
    
End Property


Public Property Get CollapseButton() As Boolean
Attribute CollapseButton.VB_Description = "Put an additional button in right corner of forms titlebar to collapse the form."
    ' Put an additional button into forms caption bar: When pressed form toggles to a smaller height (default: 0)
    
    CollapseButton = Mvar.flgBtnCollapse
    
End Property

Public Property Let CollapseButton(ByVal flgNew_CollapseButton As Boolean)
    ' Put an additional button into forms caption bar: When pressed form toggles to a smaller height (default: 0)
    
    Mvar.flgBtnCollapse = flgNew_CollapseButton
    PropertyChanged "CollapseButton"
    
End Property


Public Property Get StayOnTopButton() As Boolean
Attribute StayOnTopButton.VB_Description = "Put an additional button in right corner of forms titlebar to make form stay on top of all others."
    ' Put an additional button into forms caption bar: When pressed form floats on top of all others
    
    StayOnTopButton = Mvar.flgBtnStayOnTop
    
End Property

Public Property Let StayOnTopButton(ByVal flgNew_StayOnTopButton As Boolean)
    ' Put an additional button into forms caption bar: When pressed form floats on top of all others
    
    Mvar.flgBtnStayOnTop = flgNew_StayOnTopButton
    PropertyChanged "StayOnTopButton"
    
End Property


Public Property Get CollapseSmallSize() As Long
Attribute CollapseSmallSize.VB_Description = "When additional button is pressed we shrink the form to this value."
    ' Shrink form to this value when collapse button is pressed
    
    CollapseSmallSize = Mvar.lCollapseSmallSize
    
End Property

Public Property Let CollapseSmallSize(ByVal lNew_CollapseSmallSize As Long)
    ' Shrink form to this value when collapse button is pressed
    
    If lNew_CollapseSmallSize >= 0 Then
        Mvar.lCollapseSmallSize = lNew_CollapseSmallSize
        PropertyChanged "CollapseSmallSize"
    End If
    
End Property


Public Property Get BackgroundGradient() As enWFGradient
Attribute BackgroundGradient.VB_Description = "Select type of background gradient."
    ' Select type of gradient on forms background
    
    BackgroundGradient = Mvar.BckgrndGradient
    
End Property

Public Property Let BackgroundGradient(ByVal New_BackgroundGradient As enWFGradient)
    ' Select type of gradient on forms background
    
    Mvar.BckgrndGradient = New_BackgroundGradient
    PropertyChanged "BackgroundGradient"
    
    DrawBackgroundGradient True
    
End Property


Public Property Get BG_Width() As Long
Attribute BG_Width.VB_Description = "Width of background gradient from left form's border. 0 means:  Always full form width."
    ' Width of background gradient (0 means: Always full form width)
    
    BG_Width = ValToUsedUnit(Mvar.BGWidth)
    
End Property

Public Property Let BG_Width(ByVal New_BG_Width As Long)
    ' Width of background gradient (0 means: Always full form width)
    
    If New_BG_Width > -1 Then
        Mvar.BGWidth = UsedUnitToValue(New_BG_Width)
    End If
    PropertyChanged "BGWidth"
    
    DrawBackgroundGradient True
    
End Property


Public Property Get BG_ColorChange() As Long
Attribute BG_ColorChange.VB_Description = "Background gradient controling - IMPORTANT:  Please read docu for this!"
    ' Height or border for color change of background gradient
    
    If Mvar.BGColorChange < 0 Then
        BG_ColorChange = ValToUsedUnit(Mvar.BGColorChange)
    Else
        BG_ColorChange = Mvar.BGColorChange
    End If
    
End Property

Public Property Let BG_ColorChange(ByVal New_BG_ColorChange As Long)
    ' Height or border for color change of background gradient
    
    If New_BG_ColorChange < 0 Then
        Mvar.BGColorChange = UsedUnitToValue(New_BG_ColorChange)            ' < 0 means 'absolute' value
        
    Else
        Mvar.BGColorChange = New_BG_ColorChange                             ' Percent value (0 to 100)
        
    End If
    PropertyChanged "BGColorChange"
    
    DrawBackgroundGradient
    
End Property


Public Property Get BGColor1() As OLE_COLOR
Attribute BGColor1.VB_Description = "First background gradient color (top)"
    ' Top color background gradient
    
    BGColor1 = Mvar.BGColor1
    
End Property

Public Property Let BGColor1(ByVal New_BGColor1 As OLE_COLOR)
    ' Top color background gradient
    
    Mvar.BGColor1 = New_BGColor1
    PropertyChanged "BGColor1"
    
    DrawBackgroundGradient
        
End Property


Public Property Get BGColor2() As OLE_COLOR
Attribute BGColor2.VB_Description = "2ndt background gradient color (middle or bottom)"
    ' Middle or bottom (on two color only) color background gradient
    
    BGColor2 = Mvar.BGColor2
    
End Property

Public Property Let BGColor2(ByVal New_BGColor2 As OLE_COLOR)
    ' Middle or bottom (on two color only) color background gradient
    
    Mvar.BGColor2 = New_BGColor2
    PropertyChanged "BGColor2"
    
    DrawBackgroundGradient
    
End Property


Public Property Get BGColor3() As OLE_COLOR
Attribute BGColor3.VB_Description = "3rdt background gradient color (bottom)"
    ' Bottom (on three color only) color background gradient
    
    BGColor3 = Mvar.BGColor3
    
End Property

Public Property Let BGColor3(ByVal New_BGColor3 As OLE_COLOR)
    ' Bottom (on three color only) color background gradient
    
    Mvar.BGColor3 = New_BGColor3
    PropertyChanged "BGColor3"
    
    DrawBackgroundGradient
    
End Property


Public Property Get Unit() As enUnit
Attribute Unit.VB_Description = "Unit used for values (e.g. for 'MaxWidth' )"
    ' Unit (pixels or twips) for size values (e.g. 'max form size', 'background gradient width', ...)
    
    Unit = Mvar.Unit
    
End Property

Public Property Let Unit(ByVal New_Unit As enUnit)
    ' Unit (pixels or twips) for size values (e.g. 'max form size', 'background gradient width', ...)
    
    Mvar.Unit = New_Unit
    PropertyChanged "Unit"
    
End Property


Public Property Get FormSunken() As Boolean
    ' Activate API's Form 'Sunken' Mode - A native additional drawing mode for windows - normaly not supoorted by VB
    
    FormSunken = Mvar.flgFormSunken
    
End Property

Public Property Let FormSunken(ByVal flgNew_FormSunken As Boolean)
    ' Activate API's Form 'Sunken' Mode - A native additional drawing mode for windows - normaly not supoorted by VB
    
    Mvar.flgFormSunken = flgNew_FormSunken
    PropertyChanged "FormSunken"
    
End Property


Public Property Get MsgHandle() As Long
    ' Used to identify string messages send with WM_COPYDATA dedicated to be handled by WizzForm
    
    MsgHandle = Mvar.lMsgHandle
    
End Property

Public Property Let MsgHandle(ByVal New_MsgHandle As Long)
    ' Used to identify string  messages send with WM_COPYDATA dedicated to be handled by WizzForm.
    ' Must be <> 0 to activate message handling. Its purpose is like a password: You have to know
    ' this number to successfully send a message string which raises the Received() event.
    
    Mvar.lMsgHandle = New_MsgHandle
    
End Property


' #*#
