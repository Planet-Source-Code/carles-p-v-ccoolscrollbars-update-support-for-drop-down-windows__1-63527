VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCoolScrollbars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' Class:         cCoolScrollbars.cls
' Author:        Carles P.V. (*)
' Dependencies:  cSubclassingThunk.cls, SubclassingSink.tlb (IDE only)
' Last revision: December 10th, 2005
' Version:       1.0.2
'----------------------------------------------------------------------------------------
'
' (*) 1. cSubclassingThunk.cls
'
'        Subclassing Thunk (SuperClass V2) Project
'        Portions copyright (c) 2002 by Paul Caton <Paul_Caton@hotmail.com>
'        Portions copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'
'     2. Vital information from:
'        "Special flat Cool Scrollbars version 1.2" by James Brown
'        (http://www.catch22.net/tuts/coolscroll.asp)
'
'        This work has been decisive at time to solve following issues: how to correctly
'        get thumb size and position (see pvGetSBInfo() routine), as well as (most
'        important) *discover* that is not possible to control scrollbar paints: forget to
'        process WM_NCPAINT message.
'
'        Anyway, here, things are quite different and simpler:
'        Three static windows are created, which overlap not-client area covered by built-in
'        scrollbars. Mainly, I respond to WM_PAINT message of owner window. It's supposed
'        that any *change* related to scrollbars will raise a WM_PAINT message.
'        As mouse operates on real scrollbars, it's necessary to achieve same metrics at
'        time to paint our scrollbars. If there is any difference between real and 'virtual'
'        scrollbars metrics, you will get unexpected results.
'
'        Should say this technique might not work always (all windows).
'----------------------------------------------------------------------------------------
'
' History:
'
'   * 1.0.0: - First release.
'   * 1.0.1: - Simplified pvDrawFlatButton() routine. Old routine was used to draw glyph
'              using custom colors. Now it simply renders flat version.
'   * 1.0.2: - Support for tool-windows.
'----------------------------------------------------------------------------------------
'
' Notes:
'
'   * NOT working on 'Form' windows and 'Scrollbar' controls.
'   * Hot-tracking not supported.
'   * Thumb-tracking not supported: thumb is not painted on '.nTrackPos' positions.
'----------------------------------------------------------------------------------------
'
' Known issues:
'
'   * Problems with zOrder (parent sibling windows).
'   * NOT perfectly working on RichTextBox controls (on sizing, thumb is not correctly
'     repositioned).
'   * NOT perfectly working on ListView controls (problems with WM_TIMER (?). SetTimer API
'     returns a successful value, first WM_TIMER is sent, but something kills the timer at
'     this point).
'========================================================================================



'========================================================================================
' Class API declarations
'========================================================================================

Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const DFC_SCROLL          As Long = 3
Private Const DFCS_SCROLLUP       As Long = &H0
Private Const DFCS_SCROLLDOWN     As Long = &H1
Private Const DFCS_SCROLLSIZEGRIP As Long = &H8
Private Const DFCS_INACTIVE       As Long = &H100
Private Const DFCS_PUSHED         As Long = &H200
Private Const DFCS_FLAT           As Long = &H4000

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const COLOR_WINDOW   As Long = 5
Private Const COLOR_BTNFACE  As Long = 15
Private Const COLOR_3DSHADOW As Long = 16

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Const BLACK_BRUSH As Long = 4

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
    
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Const VK_LBUTTON As Long = &H1

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'//

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'//

Private Type PAINTSTRUCT
    hDC             As Long
    fErase          As Long
    rcPaint         As RECT
    fRestore        As Long
    fIncUpdate      As Long
    rgbReserved(32) As Byte
End Type

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long

'//

Private Const SS_CUSTOMDRAW       As Long = &HD

Private Const GWL_STYLE           As Long = -16
Private Const GWL_EXSTYLE         As Long = -20

Private Const WS_EX_TOOLWINDOW    As Long = &H80&
Private Const WS_EX_CLIENTEDGE    As Long = &H200
Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000

Private Const WS_HSCROLL          As Long = &H100000
Private Const WS_VSCROLL          As Long = &H200000
Private Const WS_BORDER           As Long = &H800000
Private Const WS_MAXIMIZE         As Long = &H1000000
Private Const WS_DISABLED         As Long = &H8000000
Private Const WS_CHILD            As Long = &H40000000

Private Const WM_ENABLE           As Long = &HA
Private Const WM_PAINT            As Long = &HF
Private Const WM_WINDOWPOSCHANGED As Long = &H47
Private Const WM_STYLECHANGED     As Long = &H7D
Private Const WM_NCLBUTTONDOWN    As Long = &HA1
Private Const WM_NCLBUTTONDBLCLK  As Long = &HA3
Private Const WM_NCRBUTTONDOWN    As Long = &HA4
Private Const WM_TIMER            As Long = &H113
Private Const WM_MOUSEMOVE        As Long = &H200
Private Const WM_CAPTURECHANGED   As Long = &H215

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'//

Private Const HWND_TOP       As Long = 0
Private Const HWND_TOPMOST   As Long = -1
Private Const SWP_NOZORDER   As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_HIDEWINDOW As Long = &H80

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'//

Private Type SCROLLINFO
    cbSize    As Long
    fMask     As Long
    nMin      As Long
    nMax      As Long
    nPage     As Long
    nPos      As Long
    nTrackPos As Long
End Type

Private Const SB_HORZ      As Long = 0
Private Const SB_VERT      As Long = 1
Private Const SIF_RANGE    As Long = &H1
Private Const SIF_PAGE     As Long = &H2
Private Const SIF_POS      As Long = &H4
Private Const SIF_TRACKPOS As Long = &H10
Private Const SIF_ALL      As Long = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS

Private Const HTHSCROLL    As Long = 6
Private Const HTVSCROLL    As Long = 7

Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long

Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3
Private Const SM_CYVSCROLL As Long = 20
Private Const SM_CXHSCROLL As Long = 21

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'//

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'//

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal Length As Long)

'//

'========================================================================================
' Class enums., variables and constants
'========================================================================================

'-- Public enums.:

Public Enum csbOnPaintPartCts
    [ppTLButton] = 0
    [ppBRButton] = 1
    [ppTLTrack] = 2
    [ppBRTrack] = 3
    [ppNullTrack] = 4
    [ppThumb] = 5
    [ppSizer] = 6
End Enum

'-- Private constants:

Private Const HT_NOTHING    As Long = 0 ' hit-test constants
Private Const HT_TLBUTTON   As Long = 1
Private Const HT_BRBUTTON   As Long = 2
Private Const HT_TLTRACK    As Long = 3
Private Const HT_BRTRACK    As Long = 4
Private Const HT_THUMB      As Long = 5

Private Const TIMERID_MOVE  As Long = 1 ' timer constants
Private Const TIMERDT_MOVE  As Long = 10

Private Const THUMBSIZE_MIN As Long = 8 ' minimum thumb size

'-- Private types:

Private Type uSCROLLBAR
    hwnd                 As Long              ' window handle
    bHasTrack            As Boolean           ' has track? (both track parts and thumb)
    lTrackLength         As Long              ' track total length
    bHasNullTrack        As Boolean           ' has null track? (buttons under minimum size)
    uRctNullTrack        As RECT              ' null track rectangle, if any
    uRctTLButton         As RECT              ' top-left button rectangle
    uRctBRButton         As RECT              ' bottom-right button rectangle
    uRctTLTrack          As RECT              ' top-left track rectangle
    uRctBRTrack          As RECT              ' bottom-right track rectangle
    uRctThumb            As RECT              ' thumb rectangle
    bTLButtonPressed     As Boolean           ' top-left button pressed?
    bBRButtonPressed     As Boolean           ' bottom-right button pressed?
    bTLTrackPressed      As Boolean           ' top-left track pressed?
    bBRTrackPressed      As Boolean           ' bottom-right track pressed?
    bThumbPressed        As Boolean           ' thumb pressed?
    lThumbOrg            As Long              ' current thumb origin (on track)
    lThumbPos            As Long              ' current thumb position (on track)
    lThumbSize           As Long              ' current thumb size
End Type

'-- Private variables:

Implements ISubclassingSink                   ' Implement Subclass interface
Private m_cSubclassW     As cSubclassingThunk ' -> Owner [W]indow
Private m_cSubclassH     As cSubclassingThunk ' -> [H]orizontal scrollbar
Private m_cSubclassV     As cSubclassingThunk ' -> [V]ertical scrollbar
Private m_cSubclassS     As cSubclassingThunk ' -> [S]izer

Private m_hWndOwner      As Long              ' owner window to apply this class to
Private m_hWndParent     As Long              ' parent window that hosts scrollbars
Private m_bOnDesktop     As Long              ' parent window is Desktop

Private m_uSB()          As uSCROLLBAR        ' type array holding scrollbars info
Private m_lSBID          As Long              ' currently processed scrollbar (hit-test)

Private m_hWndSizer      As Long              ' sizer window handle
Private m_uRctSizer      As RECT              ' sizer window rectangle

Private m_bEnabled       As Boolean           ' <enabled> flag
Private m_bLeftScrollbar As Boolean           ' <left-scrollbar> flag
Private m_bSizer         As Boolean           ' <sizer> flag

Private m_x              As Long              ' current hit-test x pos
Private m_y              As Long              ' current hit-test y pos
Private m_eHitTest       As Long              ' hit-test result

Private m_hPatternBrush  As Long              ' pattern brush

'-- R/W property variables:

Private m_bCustomDraw    As Boolean

'-- Events:

Public Event OnPaint(ByVal hDC As Long, _
                     ByVal ScrollBarID As Long, _
                     ByVal x1 As Long, ByVal y1 As Long, _
                     ByVal x2 As Long, ByVal y2 As Long, _
                     ByVal Part As csbOnPaintPartCts, _
                     ByVal Pressed As Boolean _
                     )



'========================================================================================
' Class initialization/termination
'========================================================================================

Private Sub Class_Initialize()
    
    '-- Nothing to do: all initilized on InitializeCoolSB()
End Sub

Private Sub Class_Terminate()
    
    '-- Uninitialize scrollbars
    Call Me.UninitializeCoolSB
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function InitializeCoolSB( _
                ByVal hwnd As Long, _
                Optional ByVal CustomDraw As Boolean = False _
                ) As Boolean
    
    If (m_hWndOwner = 0) Then
        
        If (IsWindow(hwnd)) Then
            
            '-- Store owner and parent window handles
            m_hWndOwner = hwnd
            m_hWndParent = GetParent(hwnd)
            m_bOnDesktop = (m_hWndParent = GetDesktopWindow())
            
            '-- Initialize subclassing thunks
            Set m_cSubclassW = New cSubclassingThunk
            Set m_cSubclassH = New cSubclassingThunk
            Set m_cSubclassV = New cSubclassingThunk
            Set m_cSubclassS = New cSubclassingThunk
                       
            '-- Initialize scrollbars related data structures and style
            ReDim m_uSB(0 To 1) As uSCROLLBAR
            m_bCustomDraw = CustomDraw
          
            '-- Create windows (both scrollbars and sizer)
            m_uSB(SB_VERT).hwnd = CreateWindowEx(-m_bOnDesktop * WS_EX_TOOLWINDOW, "Static", "", WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
            m_uSB(SB_HORZ).hwnd = CreateWindowEx(-m_bOnDesktop * WS_EX_TOOLWINDOW, "Static", "", WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
            m_hWndSizer = CreateWindowEx(-m_bOnDesktop * WS_EX_TOOLWINDOW, "Static", "", WS_CHILD Or SS_CUSTOMDRAW, 0, 0, 0, 0, m_hWndParent, 0, App.hInstance, ByVal 0)
            
            '-- Create pattern brush
            Call pvCreatePatternBrush
    
            '-- Process following owner window messages
            Call m_cSubclassW.Subclass(hwnd, Me)
            Call m_cSubclassW.AddBeforeMsgs(WM_PAINT, _
                                            WM_CAPTURECHANGED, _
                                            WM_NCLBUTTONDOWN, _
                                            WM_NCLBUTTONDBLCLK, _
                                            WM_NCRBUTTONDOWN, _
                                            WM_WINDOWPOSCHANGED, _
                                            WM_STYLECHANGED, _
                                            WM_ENABLE, _
                                            WM_TIMER)
            
            '-- Process vertical scrollbar WM_PAINT message
            Call m_cSubclassV.Subclass(m_uSB(SB_VERT).hwnd, Me)
            Call m_cSubclassV.AddBeforeMsgs(WM_PAINT)
            
            '-- Process horizontal scrollbar WM_PAINT message
            Call m_cSubclassH.Subclass(m_uSB(SB_HORZ).hwnd, Me)
            Call m_cSubclassH.AddBeforeMsgs(WM_PAINT)
            
            '-- Process sizer WM_PAINT message
            Call m_cSubclassS.Subclass(m_hWndSizer, Me)
            Call m_cSubclassS.AddBeforeMsgs(WM_PAINT)
            
            '-- Something has changed...
            Call pvOnWindowPosStyleChanged
            
            '-- Supposing all correctly done...
            InitializeCoolSB = True
        End If
    End If
End Function

Public Function UninitializeCoolSB( _
                ) As Boolean

    If (m_hWndOwner <> 0) Then
        
        '-- In case...
        Call KillTimer(m_hWndOwner, TIMERID_MOVE)
        
        '-- Stop subclassing
        Call m_cSubclassW.Unsubclass
        Call m_cSubclassH.Unsubclass
        Call m_cSubclassV.Unsubclass
        Call m_cSubclassS.Unsubclass
        
        '-- Destroy subclassing thunks
        Set m_cSubclassW = Nothing
        Set m_cSubclassH = Nothing
        Set m_cSubclassV = Nothing
        Set m_cSubclassS = Nothing
    
        '-- Destroy windows
        Call DestroyWindow(m_uSB(SB_VERT).hwnd)
        Call DestroyWindow(m_uSB(SB_HORZ).hwnd)
        Call DestroyWindow(m_hWndSizer)
        
        '-- Destroy pattern brush
        Call DeleteObject(m_hPatternBrush)
    
        '-- Something has been uninitialized
        UninitializeCoolSB = True
        m_hWndOwner = 0
    End If
End Function

Public Sub Refresh()
    
    '-- Force a complete repaint
    Call InvalidateRect(m_uSB(SB_VERT).hwnd, ByVal 0, 0)
    Call InvalidateRect(m_uSB(SB_HORZ).hwnd, ByVal 0, 0)
    Call InvalidateRect(m_hWndSizer, ByVal 0, 0)
End Sub



'========================================================================================
' Class properties
'========================================================================================

Public Property Get CustomDraw() As Boolean
    CustomDraw = m_bCustomDraw
End Property

Public Property Let CustomDraw(ByVal New_CustomDraw As Boolean)
    If (m_bCustomDraw Xor New_CustomDraw) Then
        m_bCustomDraw = New_CustomDraw
        Call Me.Refresh
    End If
End Property

Public Property Get Enabled(ByVal ScrollBarID As Long) As Boolean 'Valid only on 'OnPaint' for scrollbar parts
    Enabled = (m_bEnabled And m_uSB(ScrollBarID).lThumbSize > 0)
End Property

Public Property Get HasLeftScrollbar() As Boolean
    HasLeftScrollbar = m_bLeftScrollbar
End Property

Public Property Get HasSizer() As Boolean
    HasSizer = (m_bLeftScrollbar And m_bSizer)
End Property

Public Property Get Initialized() As Boolean
    Initialized = (m_hWndOwner <> 0)
End Property



'========================================================================================
' Messages response
'========================================================================================

Private Sub pvOnPaintScrollbar( _
            ByVal lSBID As Long, _
            ByVal hDC As Long _
            )
  
  Dim lfDir As Long
    
    lfDir = 2 * (1 - lSBID)
    
    With m_uSB(lSBID)

        If (Not m_bCustomDraw) Then
                        
            If (m_bEnabled And .lThumbSize > 0) Then
            
                '-- Buttons
                Call pvDrawFlatButton(hDC, .uRctTLButton, DFCS_SCROLLUP + lfDir, .bTLButtonPressed)
                Call pvDrawFlatButton(hDC, .uRctBRButton, DFCS_SCROLLDOWN + lfDir, .bBRButtonPressed)
                
                '-- Track + thumb
                If (.bHasTrack And .lThumbSize > 0) Then
                    
                    '-- Top-Left track part
                    If (.bTLTrackPressed) Then
                        Call FillRect(hDC, .uRctTLTrack, GetStockObject(BLACK_BRUSH))
                      Else
                        Call pvFillRectPattern(hDC, .uRctTLTrack)
                    End If
                    
                    '-- Right-Bottom track part
                    If (.bBRTrackPressed) Then
                        Call FillRect(hDC, .uRctBRTrack, GetStockObject(BLACK_BRUSH))
                      Else
                        Call pvFillRectPattern(hDC, .uRctBRTrack)
                    End If
                    
                    '-- Thumb
                    If (.lThumbSize) Then
                        If (.bThumbPressed) Then
                            Call FillRect(hDC, .uRctThumb, GetSysColorBrush(COLOR_3DSHADOW))
                          Else
                            Call DrawFrameControl(hDC, .uRctThumb, 0, DFCS_FLAT)
                        End If
                    End If
                End If
                If (.bHasNullTrack) Then
                    Call pvFillRectPattern(hDC, .uRctNullTrack)
                End If
                
              Else
                '-- Draw all disabled
                Call DrawFrameControl(hDC, .uRctTLButton, DFC_SCROLL, DFCS_SCROLLUP + lfDir Or DFCS_FLAT Or DFCS_INACTIVE)
                Call DrawFrameControl(hDC, .uRctBRButton, DFC_SCROLL, DFCS_SCROLLDOWN + lfDir Or DFCS_FLAT Or DFCS_INACTIVE)
                If (.bHasTrack) Then
                    Call pvFillRectPattern(hDC, .uRctTLTrack)
                    Call pvFillRectPattern(hDC, .uRctBRTrack)
                    If (.lThumbSize) Then
                        Call DrawFrameControl(hDC, .uRctThumb, 0, DFCS_FLAT Or DFCS_INACTIVE)
                    End If
                End If
                If (.bHasNullTrack) Then
                    Call pvFillRectPattern(hDC, .uRctNullTrack)
                End If
            End If
                
          Else
            
            With .uRctTLButton
                RaiseEvent OnPaint(hDC, lSBID, .x1, .y1, .x2, .y2, [ppTLButton], m_uSB(lSBID).bTLButtonPressed)
            End With
            With .uRctBRButton
                RaiseEvent OnPaint(hDC, lSBID, .x1, .y1, .x2, .y2, [ppBRButton], m_uSB(lSBID).bBRButtonPressed)
            End With
            If (.bHasTrack) Then
                With .uRctTLTrack
                    RaiseEvent OnPaint(hDC, lSBID, .x1, .y1, .x2, .y2, [ppTLTrack], m_uSB(lSBID).bTLTrackPressed)
                End With
                With .uRctBRTrack
                    RaiseEvent OnPaint(hDC, lSBID, .x1, .y1, .x2, .y2, [ppBRTrack], m_uSB(lSBID).bBRTrackPressed)
                End With
                If (.lThumbSize) Then
                    With .uRctThumb
                        RaiseEvent OnPaint(hDC, lSBID, .x1, .y1, .x2, .y2, [ppThumb], m_uSB(lSBID).bThumbPressed)
                    End With
                End If
            End If
            If (.bHasNullTrack) Then
                With .uRctNullTrack
                    RaiseEvent OnPaint(hDC, lSBID, .x1, .y1, .x2, .y2, [ppNullTrack], False)
                End With
            End If
        End If
    End With
End Sub

Private Sub pvOnPaintSizer( _
            ByVal hDC As Long _
            )
    
    If (Not m_bCustomDraw) Then
        If (m_bLeftScrollbar And m_bSizer) Then
            Call DrawFrameControl(hDC, m_uRctSizer, DFC_SCROLL, DFCS_SCROLLSIZEGRIP)
          Else
            Call FillRect(hDC, m_uRctSizer, GetSysColorBrush(COLOR_BTNFACE))
        End If
      Else
        With m_uRctSizer
            RaiseEvent OnPaint(hDC, -1, .x1, .y1, .x2, .y2, [ppSizer], False)
        End With
    End If
End Sub

Private Sub pvOnCaptureChanged( _
            ByVal lParam As Long _
            )
    
    If (lParam = 0) Then
        
        Call KillTimer(m_hWndOwner, TIMERID_MOVE)
        
        With m_uSB(m_lSBID)
            
            .bTLButtonPressed = False
            .bBRButtonPressed = False
            .bTLTrackPressed = False
            .bBRTrackPressed = False
            .bThumbPressed = False
        
            Call InvalidateRect(.hwnd, ByVal 0, 0)
        End With
    End If
End Sub

Private Sub pvOnMouseDown( _
            ByVal wParam As Long, _
            ByVal lParam As Long _
            )
    
    Select Case wParam
        Case HTVSCROLL
            m_lSBID = SB_VERT
        Case HTHSCROLL
            m_lSBID = SB_HORZ
        Case Else
            Exit Sub
    End Select
    
    With m_uSB(m_lSBID)
            
        If (m_bEnabled And .lThumbSize > 0) Then
        
            .bTLButtonPressed = False
            .bBRButtonPressed = False
            .bTLTrackPressed = False
            .bBRTrackPressed = False
            .bThumbPressed = False
            
            Call pvMakePoints(m_lSBID, lParam, m_x, m_y)
        
            Select Case True
                 
                 Case PtInRect(.uRctTLButton, m_x, m_y)
                     m_eHitTest = HT_TLBUTTON
                     .bTLButtonPressed = GetAsyncKeyState(VK_LBUTTON)
                 
                 Case PtInRect(.uRctBRButton, m_x, m_y)
                     m_eHitTest = HT_BRBUTTON
                     .bBRButtonPressed = GetAsyncKeyState(VK_LBUTTON)
                 
                 Case PtInRect(.uRctTLTrack, m_x, m_y)
                     m_eHitTest = HT_TLTRACK
                     .bTLTrackPressed = GetAsyncKeyState(VK_LBUTTON)
                 
                 Case PtInRect(.uRctBRTrack, m_x, m_y)
                     m_eHitTest = HT_BRTRACK
                     .bBRTrackPressed = GetAsyncKeyState(VK_LBUTTON)
                 
                 Case PtInRect(.uRctThumb, m_x, m_y)
                     m_eHitTest = HT_THUMB
                     .bThumbPressed = GetAsyncKeyState(VK_LBUTTON)
                    
                 Case Else
                     m_eHitTest = HT_NOTHING
            End Select
            
            Call InvalidateRect(.hwnd, ByVal 0, 0)
        
            If (m_eHitTest > HT_NOTHING And m_eHitTest < HT_THUMB) Then
                Call SetTimer(m_hWndOwner, TIMERID_MOVE, TIMERDT_MOVE, 0)
            End If
        End If
    End With
End Sub

Private Sub pvOnNCButtonDblClk()

    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
End Sub

Private Sub pvOnWindowPosStyleChanged()
  
  Dim lS              As Long
  Dim lSEx            As Long
  Dim bHasHorz        As Boolean
  Dim bHasVert        As Boolean
  Dim uRctParent      As RECT
  Dim uRctOwner       As RECT
  Dim uRctOwnerClient As RECT
  Dim uPt             As POINTAPI
  Dim lEdge           As Long
  Dim hWndInsertAfter As Long
  
    '-- Get window (owner) style and extended style bits
    lS = GetWindowLong(m_hWndOwner, GWL_STYLE)
    lSEx = GetWindowLong(m_hWndOwner, GWL_EXSTYLE)
    
    '-- Is window enabled?
    m_bEnabled = Not ((lS And WS_DISABLED) = WS_DISABLED)
    
    '-- Get owner window and client rectangles
    Call GetWindowRect(m_hWndOwner, uRctOwner)
    Call GetClientRect(m_hWndOwner, uRctOwnerClient)
    Call CopyMemory(uPt, uRctOwner, LenB(uPt))
    Call ScreenToClient(m_hWndParent, uPt)
    
    '-- Left vertical scrollbar?
    m_bLeftScrollbar = ((lSEx And WS_EX_LEFTSCROLLBAR) = 0)
    
    '-- Get edge length
    lEdge = -1 * ((lS And WS_BORDER) = WS_BORDER)
    lEdge = lEdge + -2 * ((lSEx And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE)
    
    '-- Has scrollbars?
    With uRctOwner
        bHasHorz = ((lS And WS_HSCROLL) = WS_HSCROLL) And (.y2 - .y1 - 2 * lEdge > GetSystemMetrics(SM_CYHSCROLL))
        bHasVert = ((lS And WS_VSCROLL) = WS_VSCROLL) And (.x2 - .x1 - 2 * lEdge > GetSystemMetrics(SM_CXVSCROLL))
    End With
    
    '-- Show/hide and resize windows
    With uRctOwnerClient
        
        '-- Owned by Desktop
        If (m_bOnDesktop) Then
            hWndInsertAfter = HWND_TOPMOST
          Else
            hWndInsertAfter = HWND_TOP
        End If
        
        '-- It has vertical scroll bar...
        If (bHasVert) Then
            Call SetWindowPos(m_uSB(SB_VERT).hwnd, hWndInsertAfter, uPt.x + lEdge + .x2 * -m_bLeftScrollbar, uPt.y + lEdge, GetSystemMetrics(SM_CXVSCROLL), .y2, SWP_SHOWWINDOW)
            Call pvSizeButtons(SB_VERT)
          Else
            Call SetWindowPos(m_uSB(SB_VERT).hwnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_HIDEWINDOW)
        End If
        '-- It has horizontal scroll bar...
        If (bHasHorz) Then
            Call SetWindowPos(m_uSB(SB_HORZ).hwnd, hWndInsertAfter, uPt.x + lEdge + GetSystemMetrics(SM_CXVSCROLL) * -(Not m_bLeftScrollbar And bHasVert), uPt.y + .y2 + lEdge, .x2, GetSystemMetrics(SM_CYHSCROLL), SWP_SHOWWINDOW)
            Call pvSizeButtons(SB_HORZ)
          Else
            Call SetWindowPos(m_uSB(SB_HORZ).hwnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_HIDEWINDOW)
        End If
        '-- It has both: add 'sizer window'
        If (bHasVert And bHasHorz) Then
            Call SetWindowPos(m_hWndSizer, hWndInsertAfter, uPt.x + lEdge + .x2 * -m_bLeftScrollbar, uPt.y + lEdge + .y2, GetSystemMetrics(SM_CXVSCROLL), GetSystemMetrics(SM_CYHSCROLL), SWP_SHOWWINDOW)
          Else
            Call SetWindowPos(m_hWndSizer, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_HIDEWINDOW)
        End If
    End With
    
    '-- Check if 'sizer' is needed
    If (bHasVert And bHasHorz) Then
 
        Call GetClientRect(m_hWndSizer, m_uRctSizer)
        Call OffsetRect(uRctOwner, -uRctOwner.x1, -uRctOwner.y1)
        Call GetClientRect(m_hWndParent, uRctParent)
        lS = GetWindowLong(m_hWndParent, GWL_STYLE)

        m_bSizer = (EqualRect(uRctOwner, uRctParent) <> 0 And ((lS And WS_MAXIMIZE) = 0))
      Else
        m_bSizer = False
    End If
    
    '-- Force repaint
    Call Me.Refresh
End Sub

Private Sub pvOnEnable( _
            ByVal wParam As Long _
            )

    m_bEnabled = CBool(wParam)
    Call Me.Refresh
End Sub

Private Sub pvOnTimer( _
            ByVal wParam As Long _
            )
  
  Dim uPt      As POINTAPI
  Dim bPressed As Boolean
  
    Select Case wParam

        Case TIMERID_MOVE
            
            With m_uSB(m_lSBID)
                
                Call GetCursorPos(uPt)
                Call ScreenToClient(.hwnd, uPt)
                
                Select Case m_eHitTest
                    
                    Case HT_TLBUTTON
                        bPressed = (PtInRect(.uRctTLButton, uPt.x, uPt.y) <> 0)
                        If (.bTLButtonPressed Xor bPressed) Then
                            .bTLButtonPressed = bPressed
                            Call InvalidateRect(.hwnd, ByVal 0, 0)
                        End If
                    
                    Case HT_BRBUTTON
                        bPressed = (PtInRect(.uRctBRButton, uPt.x, uPt.y) <> 0)
                        If (.bBRButtonPressed Xor bPressed) Then
                            .bBRButtonPressed = bPressed
                            Call InvalidateRect(.hwnd, ByVal 0, 0)
                        End If
                    
                    Case HT_TLTRACK
                        bPressed = (PtInRect(.uRctTLTrack, uPt.x, uPt.y) <> 0)
                        If (.bTLTrackPressed Xor bPressed) Then
                            .bTLTrackPressed = bPressed
                            Call InvalidateRect(.hwnd, ByVal 0, 0)
                        End If
                    
                    Case HT_BRTRACK
                        bPressed = (PtInRect(.uRctBRTrack, uPt.x, uPt.y) <> 0)
                        If (.bBRTrackPressed Xor bPressed) Then
                            .bBRTrackPressed = bPressed
                            Call InvalidateRect(.hwnd, ByVal 0, 0)
                        End If
               End Select
           End With
    End Select
End Sub



'========================================================================================
' Private
'========================================================================================

'----------------------------------------------------------------------------------------
' Sizing
'----------------------------------------------------------------------------------------

Private Sub pvGetSBInfo( _
            ByVal lSBID As Long _
            )

  Dim uSBI As SCROLLINFO
  Dim lAbs As Long
  Dim lPos As Long
  
    uSBI.fMask = SIF_ALL
    uSBI.cbSize = LenB(uSBI)
    
    Call GetScrollInfo(m_hWndOwner, lSBID, uSBI)
    
    With m_uSB(lSBID)
        
        If (uSBI.nPage <= uSBI.nMax) Then
            lAbs = uSBI.nMax - uSBI.nMin + 1
            .lThumbSize = MulDiv(uSBI.nPage, .lTrackLength, lAbs)
            If (.lThumbSize < THUMBSIZE_MIN) Then
                .lThumbSize = THUMBSIZE_MIN
            End If
            If (m_eHitTest = HT_THUMB) Then
                lPos = uSBI.nTrackPos
              Else
                lPos = uSBI.nPos
            End If
            .lThumbPos = .lThumbOrg + MulDiv(lPos - uSBI.nMin, .lTrackLength - .lThumbSize, lAbs - uSBI.nPage)
          Else
            .lThumbSize = 0
            .lThumbPos = .lThumbOrg
        End If
    End With
End Sub

Private Sub pvSizeButtons( _
            ByVal lSBID As Long _
            )
 
  Dim uRct        As RECT
  Dim lButtonSize As Long
    
    With m_uSB(lSBID)
        
        .bHasTrack = False
        .bHasNullTrack = False
        
        If (lSBID = SB_VERT) Then
    
            Call GetClientRect(.hwnd, uRct)
            lButtonSize = GetSystemMetrics(SM_CYVSCROLL)
            
            If (2 * lButtonSize + THUMBSIZE_MIN > uRct.y2) Then
                If (2 * lButtonSize < uRct.y2) Then
                    Call SetRect(.uRctTLButton, 0, 0, uRct.x2, lButtonSize)
                    Call SetRect(.uRctBRButton, 0, uRct.y2 - lButtonSize, uRct.x2, uRct.y2)
                    .bHasNullTrack = True
                    Call SetRect(.uRctNullTrack, 0, lButtonSize, uRct.x2, uRct.y2 - lButtonSize)
                  Else
                    Call SetRect(.uRctTLButton, 0, 0, uRct.x2, uRct.y2 \ 2)
                    Call SetRect(.uRctBRButton, 0, uRct.y2 \ 2 + (uRct.y2 Mod 2), uRct.x2, uRct.y2)
                    .bHasNullTrack = CBool(uRct.y2 Mod 2)
                    If (.bHasNullTrack) Then
                        Call SetRect(.uRctNullTrack, 0, uRct.y2 \ 2, uRct.x2, uRct.y2 \ 2 + 1)
                    End If
                End If
              Else
                .bHasTrack = True
                Call SetRect(.uRctTLButton, 0, 0, uRct.x2, lButtonSize)
                Call SetRect(.uRctBRButton, 0, uRct.y2 - lButtonSize, uRct.x2, uRct.y2)
            End If
            .lThumbOrg = .uRctTLButton.y2
            .lTrackLength = .uRctBRButton.y1 - .uRctTLButton.y2
            
          Else
            
            Call GetClientRect(.hwnd, uRct)
            lButtonSize = GetSystemMetrics(SM_CXHSCROLL)
                
            If (2 * lButtonSize + THUMBSIZE_MIN > uRct.x2) Then
                If (2 * lButtonSize < uRct.x2) Then
                    Call SetRect(.uRctTLButton, 0, 0, lButtonSize, uRct.y2)
                    Call SetRect(.uRctBRButton, uRct.x2 - lButtonSize, 0, uRct.x2, uRct.y2)
                    .bHasNullTrack = True
                    Call SetRect(.uRctNullTrack, lButtonSize, 0, uRct.x2 - lButtonSize, uRct.y2)
                  Else
                    Call SetRect(.uRctTLButton, 0, 0, uRct.x2 \ 2, uRct.y2)
                    Call SetRect(.uRctBRButton, uRct.x2 \ 2 + (uRct.x2 Mod 2), 0, uRct.x2, uRct.y2)
                    .bHasNullTrack = CBool(uRct.x2 Mod 2)
                    If (.bHasNullTrack) Then
                        Call SetRect(.uRctNullTrack, uRct.x2 \ 2, 0, uRct.x2 \ 2 + 1, uRct.y2)
                    End If
                End If
              Else
                .bHasTrack = True
                Call SetRect(.uRctTLButton, 0, 0, lButtonSize, uRct.y2)
                Call SetRect(.uRctBRButton, uRct.x2 - lButtonSize, 0, uRct.x2, uRct.y2)
            End If
            .lThumbOrg = .uRctTLButton.x2
            .lTrackLength = .uRctBRButton.x1 - .uRctTLButton.x2
        End If
    End With
End Sub

Private Sub pvSizeTrack( _
            ByVal lSBID As Long _
            )
    
    With m_uSB(lSBID)
    
        If (lSBID = SB_VERT) Then
            Call SetRect(.uRctTLTrack, 0, .uRctTLButton.y2, .uRctTLButton.x2, .lThumbPos)
            Call SetRect(.uRctBRTrack, 0, .lThumbPos + .lThumbSize, .uRctBRButton.x2, .uRctBRButton.y1)
            Call SetRect(.uRctThumb, 0, .lThumbPos, .uRctBRButton.x2, .lThumbPos + .lThumbSize)
          Else
            Call SetRect(.uRctTLTrack, .uRctTLButton.x2, 0, .lThumbPos, .uRctTLButton.y2)
            Call SetRect(.uRctBRTrack, .lThumbPos + .lThumbSize, 0, .uRctBRButton.x1, .uRctBRButton.y2)
            Call SetRect(.uRctThumb, .lThumbPos, 0, .lThumbPos + .lThumbSize, .uRctBRButton.y2)
        End If
    End With
End Sub

'----------------------------------------------------------------------------------------
' Painting
'----------------------------------------------------------------------------------------

Private Sub pvDrawFlatButton( _
            ByVal hDC As Long, _
            uRct As RECT, _
            ByVal lfArrowDirection As Long, _
            ByVal bPressed As Boolean _
            )

  Dim uRctMem As RECT
  
  Dim hDCMem  As Long
  Dim hBmp    As Long
  Dim hBmpOld As Long
        
    With uRct
    
        '-- Buffer DC
        hDCMem = CreateCompatibleDC(hDC)
        hBmp = CreateCompatibleBitmap(hDC, .x2 - .x1, .y2 - .y1)
        hBmpOld = SelectObject(hDCMem, hBmp)
        
        '-- Draw frame normally
        Call CopyRect(uRctMem, uRct)
        Call OffsetRect(uRctMem, -.x1, -.y1)
        Call DrawFrameControl(hDCMem, uRctMem, DFC_SCROLL, DFCS_FLAT Or lfArrowDirection)
        
        '-- Paint to dest. DC
        If (bPressed) Then
            '-- Invert
            Call InvertRect(hDCMem, uRctMem)
        End If
        Call BitBlt(hDC, .x1, .y1, .x2 - .x1, .y2 - .y1, hDCMem, 0, 0, vbSrcCopy)
    End With
        
    '-- Clean up
    Call DeleteObject(SelectObject(hDCMem, hBmpOld))
    Call DeleteDC(hDCMem)
End Sub

Private Sub pvFillRectPattern( _
            ByVal hDC As Long, _
            lpRect As RECT _
            )
    
 Dim clrBk   As Long
 Dim clrText As Long
    
    clrBk = SetBkColor(hDC, GetSysColor(COLOR_WINDOW))
    clrText = SetTextColor(hDC, GetSysColor(COLOR_BTNFACE))
    
    Call FillRect(hDC, lpRect, m_hPatternBrush)
    
    Call SetBkColor(hDC, clrBk)
    Call SetTextColor(hDC, clrText)
End Sub

'----------------------------------------------------------------------------------------
' Misc.
'----------------------------------------------------------------------------------------

Private Sub pvCreatePatternBrush()

  Dim hBitmap As Long
    
    '-- Create brush from bitmap
    hBitmap = CreateBitmap(8, 8, 1, 1, ByVal "ªªUUªªUUªªUUªªUU")
    m_hPatternBrush = CreatePatternBrush(hBitmap)
    Call DeleteObject(hBitmap)
End Sub

Private Sub pvMakePoints( _
            ByVal lSBID As Long, _
            ByVal lPoint As Long, _
            x As Long, _
            y As Long _
            )
               
  Dim uRct As RECT
    
    If (lPoint And &H8000&) Then
        x = &H8000 Or (lPoint And &H7FFF&)
      Else
        x = lPoint And &HFFFF&
    End If
    
    If (lPoint And &H80000000) Then
        y = (lPoint \ &H10000) - 1
      Else
        y = lPoint \ &H10000
    End If
    
    '-- Translate to not-client
    Call GetWindowRect(m_uSB(lSBID).hwnd, uRct)
    x = x - uRct.x1
    y = y - uRct.y1
End Sub



'========================================================================================
' Subclass implementation
'========================================================================================

Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)

  Dim uPS As PAINTSTRUCT
  
    Select Case hwnd
      
        Case m_hWndOwner
      
            Select Case uMsg
            
                Case WM_PAINT
                    Call Me.Refresh

                Case WM_CAPTURECHANGED
                    Call pvOnCaptureChanged(lParam)
                
                Case WM_NCLBUTTONDOWN
                    Call pvOnMouseDown(wParam, lParam)
                    
                Case WM_NCRBUTTONDOWN
                    bHandled = Not m_bEnabled
                    
                Case WM_NCLBUTTONDBLCLK
                    Call pvOnNCButtonDblClk
                    bHandled = True
                    
                Case WM_WINDOWPOSCHANGED, WM_STYLECHANGED
                    Call pvOnWindowPosStyleChanged
                   
                Case WM_ENABLE
                    Call pvOnEnable(wParam)
                
                Case WM_TIMER
                    Call pvOnTimer(wParam)
           End Select
                
        Case m_uSB(SB_VERT).hwnd
        
            If (uMsg = WM_PAINT) Then
                Call pvGetSBInfo(SB_VERT)
                Call pvSizeTrack(SB_VERT)
                Call BeginPaint(m_uSB(SB_VERT).hwnd, uPS)
                Call pvOnPaintScrollbar(SB_VERT, uPS.hDC)
                Call EndPaint(m_uSB(SB_VERT).hwnd, uPS)
            End If
        
        Case m_uSB(SB_HORZ).hwnd
        
            If (uMsg = WM_PAINT) Then
                Call pvGetSBInfo(SB_HORZ)
                Call pvSizeTrack(SB_HORZ)
                Call BeginPaint(m_uSB(SB_HORZ).hwnd, uPS)
                Call pvOnPaintScrollbar(SB_HORZ, uPS.hDC)
                Call EndPaint(m_uSB(SB_HORZ).hwnd, uPS)
            End If
            
        Case m_hWndSizer
            
            If (uMsg = WM_PAINT) Then
                Call BeginPaint(m_hWndSizer, uPS)
                Call pvOnPaintSizer(uPS.hDC)
                Call EndPaint(m_hWndSizer, uPS)
            End If
    End Select
End Sub

Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

End Sub
