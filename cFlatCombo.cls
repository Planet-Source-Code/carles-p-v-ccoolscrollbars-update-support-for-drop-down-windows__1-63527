VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlatCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' Class:         cFlatCombo.cls
' Author:        Carles P.V. - 2005 (*)
' Dependencies:  cSubclassingThunk.cls, SubclassingSink.tlb (IDE only)
' Last revision: December 14th, 2005
' Version:       1.0.3
'----------------------------------------------------------------------------------------
'
' (*) cSubclassingThunk.cls
'
'     Subclassing Thunk (SuperClass V2) Project
'     Portions copyright (c) 2002 by Paul Caton <Paul_Caton@hotmail.com>
'     Portions copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'----------------------------------------------------------------------------------------
'
' History:
'
'   * 1.0.0: - First release.
'   * 1.0.1: - Fixed paint on WM_SETFOCUS/WM_KILLFOCUS
'   * 1.0.2: - Sorry: removed now unnecessary processing of WM_LBUTTONDOWN, WM_LBUTTONUP
'              and WM_MOUSEMOVE messages. Previously (old code) needed because of button
'              rectangle did not go with painted button rectangle.
'   * 1.0.3: - bEnabled state extracted from window style bits instead of '.Enabled' control
'              property. '.Enabled' property returned False when 'popup-modal' window was
'              activated.
'========================================================================================

Option Explicit

'========================================================================================
' API declarations
'========================================================================================

'== Structures

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type PAINTSTRUCT
    hDC             As Long
    fErase          As Long
    rcPaint         As RECT
    fRestore        As Long
    fIncUpdate      As Long
    rgbReserved(32) As Byte
End Type

'== Message constants

Private Const WM_SETFOCUS          As Long = &H7
Private Const WM_KILLFOCUS         As Long = &H8
Private Const WM_PAINT             As Long = &HF
Private Const WM_ERASEBKGND        As Long = &H14
Private Const WM_WINDOWPOSCHANGING As Long = &H46
Private Const WM_COMMAND           As Long = &H111
Private Const WM_CTLCOLORLISTBOX   As Long = &H134

Private Const CB_GETDROPPEDSTATE   As Long = &H157

'== Misc.

Private Const GWL_STYLE   As Long = -16
Private Const WS_DISABLED As Long = &H8000000
Private Const WS_VISIBLE  As Long = &H10000000

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const BDR_SUNKEN As Long = &HA
Private Const BF_RECT    As Long = &HF
Private Const BF_MONO    As Long = &H8000

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Const DFC_SCROLL      As Long = 3
Private Const DFCS_SCROLLDOWN As Long = &H1
Private Const DFCS_INACTIVE   As Long = &H100
Private Const DFCS_PUSHED     As Long = &H200
Private Const DFCS_FLAT       As Long = &H4000

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long

Private Const DT_VCENTER    As Long = &H4
Private Const DT_SINGLELINE As Long = &H20

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const COLOR_WINDOW        As Long = 5
Private Const COLOR_WINDOWTEXT    As Long = 8
Private Const COLOR_HIGHLIGHT     As Long = 13
Private Const COLOR_HIGHLIGHTTEXT As Long = 14
Private Const COLOR_GRAYTEXT      As Long = 17

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Const VK_LBUTTON As Long = &H1

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const TRANSPARENT As Long = 1

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Private Const SM_CXVSCROLL As Long = 2

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long



'========================================================================================
' Class enums., variables, constants...
'========================================================================================

'-- Interfaces:
Implements ISubclassingSink

'-- Private variables:
Private m_oCombo      As Object            ' private reference to subclassed ComboBox
Private m_oSubclass   As cSubclassingThunk ' subclassing class
Private m_oScrollbars As cCoolScrollbars   ' cool-scrollbars class
Private m_bFocused    As Boolean

'========================================================================================
' Methods
'========================================================================================

Public Sub Initialize(oComboBox As Object)

    If (VBA.TypeName(oComboBox) = "ComboBox") Then
        
        Set m_oCombo = oComboBox
        
        Set m_oSubclass = New cSubclassingThunk
        Call m_oSubclass.Subclass(oComboBox.hwnd, Me)
        Call m_oSubclass.AddBeforeMsgs(WM_PAINT, WM_ERASEBKGND, WM_SETFOCUS, WM_KILLFOCUS, WM_COMMAND)
        Call m_oSubclass.AddAfterMsgs(WM_SETFOCUS, WM_KILLFOCUS, WM_COMMAND, WM_CTLCOLORLISTBOX, WM_WINDOWPOSCHANGING)
    End If
End Sub

Public Sub Uninitialize()

    If (Not m_oSubclass Is Nothing) Then
        Call m_oSubclass.Unsubclass
        Set m_oSubclass = Nothing
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvDrawComboBox(ByVal hDC As Long)
  
  Dim rcCombo  As RECT
  Dim rcButton As RECT
  Dim rcText   As RECT
  Dim pt       As POINTAPI
  
  Dim hDCMem   As Long
  Dim hBmp     As Long
  Dim hBmpOld  As Long
  Dim oFont    As IFont
  Dim hFontOld As Long
  Dim bDropped As Boolean
  Dim bPressed As Boolean
  Dim bEnabled As Boolean
  
    With m_oCombo
        
        '-- Get combo window rect.:
        '   Combo does not have *real* edge.
        Call GetClientRect(.hwnd, rcCombo)
        
        '-- Calc. button rect.
        LSet rcButton = rcCombo
        Call InflateRect(rcButton, -2, -2)
        rcButton.x1 = rcButton.x2 - GetSystemMetrics(SM_CXVSCROLL)
       
        '-- Create buffer DC
        hDCMem = CreateCompatibleDC(hDC)
        hBmp = CreateCompatibleBitmap(hDC, rcCombo.x2, rcCombo.y2)
        Call SelectObject(hDCMem, hBmp)
        
        '-- Prepare font (thanks Vlad for IFont tip)
        Set oFont = .Font
        hFontOld = SelectObject(hDCMem, oFont.hFont)
        Call SetBkMode(hDCMem, TRANSPARENT)
        
        '-- Get state values
        Call GetCursorPos(pt)
        Call ScreenToClient(.hwnd, pt)
        bDropped = CBool(SendMessage(.hwnd, CB_GETDROPPEDSTATE, 0, ByVal 0))
        bPressed = CBool(PtInRect(rcButton, pt.x, pt.y)) And CBool(GetAsyncKeyState(VK_LBUTTON))
        bEnabled = Not ((GetWindowLong(.hwnd, GWL_STYLE) And WS_DISABLED) = WS_DISABLED)
        
        '-- Erase background and draw edge
        Call FillRect(hDCMem, rcCombo, GetSysColorBrush(COLOR_WINDOW))
        Call DrawEdge(hDCMem, rcCombo, BDR_SUNKEN, BF_RECT + BF_MONO)
        
        '-- Draw button
        If (.Style <> 1) Then ' not [Simple Combo]
            Call DrawFrameControl(hDCMem, rcButton, DFC_SCROLL, DFCS_SCROLLDOWN + DFCS_FLAT + DFCS_PUSHED * -(bPressed And bDropped) + DFCS_INACTIVE * -(Not bEnabled))
        End If
        
        '-- Draw text
        If (.Style = 2) Then ' [Dropdown List]
            
            '-- Calc. text rect.
            LSet rcText = rcCombo
            Call InflateRect(rcText, -2, -2)
            rcText.x2 = rcButton.x1 - 1
            rcText.x1 = 2
            
            If (Not bEnabled) Then ' disabled
                
                '-- Draw text disabled
                Call InflateRect(rcText, -2, 0)
                Call SetTextColor(hDCMem, GetSysColor(COLOR_GRAYTEXT))
                Call DrawText(hDCMem, .Text, -1, rcText, DT_SINGLELINE + DT_VCENTER)
              
              Else
                
                If (m_bFocused And Not bDropped) Then ' focused
                    
                    '-- Selected background
                    Call FillRect(hDCMem, rcText, GetSysColorBrush(COLOR_HIGHLIGHT))
                    Call InflateRect(rcText, -2, 0)
                    
                    '-- Selected text
                    Call SetTextColor(hDCMem, GetSysColor(COLOR_HIGHLIGHTTEXT))
                    Call DrawText(hDCMem, .Text, -1, rcText, DT_SINGLELINE + DT_VCENTER)
                    
                    '-- Focus rectangle
                    Call InflateRect(rcText, 2, 0)
                    Call SetTextColor(hDCMem, 0)
                    Call DrawFocusRect(hDCMem, rcText)
                  
                  Else ' not focused
                    
                    '-- Normal text
                    Call InflateRect(rcText, -2, 0)
                    Call SetTextColor(hDCMem, GetSysColor(COLOR_WINDOWTEXT))
                    Call DrawText(hDCMem, .Text, -1, rcText, DT_SINGLELINE + DT_VCENTER)
                End If
            End If
        End If
        
        '-- BitBlt to combo DC
        Call BitBlt(hDC, 0, 0, rcCombo.x2, rcCombo.y2, hDCMem, 0, 0, vbSrcCopy)
        
        '-- Clean up
        Call SelectObject(hDCMem, hFontOld)
        Call SelectObject(hDCMem, hBmpOld)
        Call DeleteObject(hBmp)
        Call DeleteDC(hDCMem)
    End With
End Sub

Private Sub pvSetVisible(ByVal hwnd As Long, ByVal bVisible As Boolean)
    
  Dim lS As Long
  
    '-- Enable/disable window visibility
    lS = GetWindowLong(hwnd, GWL_STYLE)
    If (bVisible) Then
        lS = lS Or WS_VISIBLE
      Else
        lS = lS And Not WS_VISIBLE
    End If
    Call SetWindowLong(hwnd, GWL_STYLE, lS)
End Sub

'========================================================================================
' Message response
'========================================================================================

Private Sub ISubclassingSink_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
                          
 Dim uPS As PAINTSTRUCT

    Select Case uMsg
    
        Case WM_ERASEBKGND
            bHandled = True: lReturn = 1
            
        Case WM_PAINT
            Call InvalidateRect(hwnd, ByVal 0, 0)
            Call pvDrawComboBox(BeginPaint(hwnd, uPS))
            Call EndPaint(hwnd, uPS)
            bHandled = True: lReturn = 0
                    
        Case WM_SETFOCUS, WM_KILLFOCUS
            m_bFocused = (uMsg = WM_SETFOCUS)
            Call pvSetVisible(hwnd, False)
        
        Case WM_COMMAND
            Call pvSetVisible(hwnd, False)
    End Select
End Sub

Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
                          
    Select Case uMsg
    
        Case WM_SETFOCUS, WM_KILLFOCUS
            Call pvSetVisible(hwnd, True)
        
        Case WM_COMMAND
            Call pvSetVisible(hwnd, True)
            
        Case WM_CTLCOLORLISTBOX
            If (m_oScrollbars Is Nothing) Then
                Set m_oScrollbars = New cCoolScrollbars
                Call m_oScrollbars.InitializeCoolSB(lParam)
            End If
            
        Case WM_WINDOWPOSCHANGING
            If (Not m_oScrollbars Is Nothing) Then
                Call m_oScrollbars.UninitializeCoolSB
                Set m_oScrollbars = Nothing
            End If
    End Select
End Sub
