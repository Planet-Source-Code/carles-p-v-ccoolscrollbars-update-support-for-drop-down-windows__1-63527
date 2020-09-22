VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlatImageCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' Class:         cFlatImageCombo.cls
' Author:        Carles P.V. - 2005 (*)
' Dependencies:  cSubclassingThunk.cls, SubclassingSink.tlb (IDE only)
' Last revision: December 14th, 2005
' Version:       1.0.2
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
'   * 1.0.1: - Sorry: removed now unnecessary processing of WM_LBUTTONDOWN, WM_LBUTTONUP
'              and WM_MOUSEMOVE messages. Previously (old code) needed because of button
'              rectangle did not go with painted button rectangle.
'   * 1.0.2: - bEnabled state extracted from window style bits instead of '.Enabled' control
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
Private Const WM_CTLCOLORLISTBOX   As Long = &H134
Private Const WM_PRINTCLIENT       As Long = &H318

Private Const CB_GETDROPPEDSTATE   As Long = &H157

'== Misc.

Private Const GWL_STYLE   As Long = -16
Private Const WS_DISABLED As Long = &H8000000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GW_CHILD As Long = 5

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
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

Private Const COLOR_WINDOW        As Long = 5

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Const VK_LBUTTON As Long = &H1

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SM_CXVSCROLL As Long = 2

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long



'========================================================================================
' Class enums., variables, constants...
'========================================================================================

'-- Interfaces:
Implements ISubclassingSink

'-- Private variables:
Private m_oImageCombo    As Object            ' private reference to subclassed ImageCombo
Private m_oSubclass      As cSubclassingThunk ' subclassing class
Private m_oSubclassChild As cSubclassingThunk ' subclassing class (child)
Private m_oScrollbars    As cCoolScrollbars   ' cool-scrollbars class
Private m_bFocused       As Boolean

'========================================================================================
' Methods
'========================================================================================

Public Sub Initialize(oImageCombo As Object)
  
    If (VBA.TypeName(oImageCombo) = "ImageCombo") Then
        
        Set m_oImageCombo = oImageCombo
        
        Set m_oSubclass = New cSubclassingThunk
        Call m_oSubclass.Subclass(oImageCombo.hwnd, Me)
        Call m_oSubclass.AddBeforeMsgs(WM_ERASEBKGND)
        
        Set m_oSubclassChild = New cSubclassingThunk
        Call m_oSubclassChild.Subclass(GetWindow(oImageCombo.hwnd, GW_CHILD), Me)
        Call m_oSubclassChild.AddBeforeMsgs(WM_PAINT, WM_ERASEBKGND, WM_PRINTCLIENT, WM_SETFOCUS, WM_KILLFOCUS)
        Call m_oSubclassChild.AddAfterMsgs(WM_CTLCOLORLISTBOX, WM_WINDOWPOSCHANGING)
    End If
End Sub

Public Sub Uninitialize()

    If (Not m_oSubclass Is Nothing) Then
        Call m_oSubclass.Unsubclass
        Set m_oSubclass = Nothing
    End If
    If (Not m_oSubclassChild Is Nothing) Then
        Call m_oSubclassChild.Unsubclass
        Set m_oSubclassChild = Nothing
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvDrawImageCombo(ByVal hDC As Long)
  
  Dim rcCombo  As RECT
  Dim rcButton As RECT
  Dim pt       As POINTAPI
  
  Dim hDCMem   As Long
  Dim hBmp     As Long
  Dim hBmpOld  As Long
  Dim bDropped As Boolean
  Dim bPressed As Boolean
  Dim bEnabled As Boolean
  Dim lx       As Long
  Dim ly       As Long
  
    With m_oImageCombo
        
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
        Call DrawFrameControl(hDCMem, rcButton, DFC_SCROLL, DFCS_SCROLLDOWN + DFCS_FLAT + DFCS_PUSHED * -(bPressed And bDropped) + DFCS_INACTIVE * -(Not bEnabled))
        
        '-- Draw icon
        If (Not .SelectedItem Is Nothing) Then
            If (Not .ImageList Is Nothing) Then
                lx = 2 + 1
                ly = (rcCombo.y2 - rcCombo.y1 - .ImageList.ImageHeight) / 2
                lx = .ImageList.Parent.ScaleX(lx, vbPixels, .ImageList.Parent.ScaleMode)
                ly = .ImageList.Parent.ScaleY(ly, vbPixels, .ImageList.Parent.ScaleMode)
                Select Case True
                    Case .SelectedItem.SelImage <> 0
                        Call .ImageList.ListImages(.SelectedItem.SelImage).Draw(hDCMem, lx, ly, 1 + 2 * -(m_bFocused And Not bDropped))
                    Case .SelectedItem.Image <> 0
                        Call .ImageList.ListImages(.SelectedItem.Image).Draw(hDCMem, lx, ly, 1 + 2 * -(m_bFocused And Not bDropped))
                End Select
            End If
        End If
        
        '-- BitBlt to combo DC
        Call BitBlt(hDC, 0, 0, rcCombo.x2, rcCombo.y2, hDCMem, 0, 0, vbSrcCopy)
        
        '-- Clean up
        Call SelectObject(hDCMem, hBmpOld)
        Call DeleteObject(hBmp)
        Call DeleteDC(hDCMem)
    End With
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
            Call pvDrawImageCombo(BeginPaint(hwnd, uPS))
            Call EndPaint(hwnd, uPS)
            bHandled = True: lReturn = 0

        Case WM_PRINTCLIENT
            bHandled = True: lReturn = 0
                    
        Case WM_SETFOCUS, WM_KILLFOCUS
            m_bFocused = (uMsg = WM_SETFOCUS)
    End Select
End Sub

Private Sub ISubclassingSink_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
                          
    Select Case uMsg
        
        Case WM_CTLCOLORLISTBOX
            If (m_oScrollbars Is Nothing) Then
                Set m_oScrollbars = New cCoolScrollbars
                Call m_oScrollbars.InitializeCoolSB(lParam)
                m_bFocused = True
            End If
            
        Case WM_WINDOWPOSCHANGING
            If (Not m_oScrollbars Is Nothing) Then
                Call m_oScrollbars.UninitializeCoolSB
                Set m_oScrollbars = Nothing
            End If
    End Select
End Sub
