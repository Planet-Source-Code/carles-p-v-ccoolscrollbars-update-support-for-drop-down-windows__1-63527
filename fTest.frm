VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Demo"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   255
      Width           =   2400
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnuTestTop 
      Caption         =   "&Test"
      Begin VB.Menu mnuTest 
         Caption         =   "&Initialize"
         Index           =   0
      End
      Begin VB.Menu mnuTest 
         Caption         =   "&Uninitialize"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuTest 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTest 
         Caption         =   "&Custom draw"
         Index           =   3
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oSB As cCoolScrollbars ' WithEvents: only if you want yo make use of custom-draw
Attribute m_oSB.VB_VarHelpID = -1
Private Const clrMYCOLOR As Long = &HBF9F7F



Private Sub Form_Load()

  Dim hFile As Long
  Dim sFile As String

    sFile = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "cCoolScrollbars.cls"
    hFile = FreeFile()
    Open sFile For Input As hFile
        Text1.Text = Input$(LOF(hFile), hFile)
    Close hFile
    
    Set m_oSB = New cCoolScrollbars
End Sub

'//

Private Sub mnuFile_Click(Index As Integer)
    Call Unload(Me)
End Sub

Private Sub mnuTest_Click(Index As Integer)

    Select Case Index
        
        Case 0 '-- Initialize cool scrollbars for passed window handle
            Call m_oSB.InitializeCoolSB(hwnd:=Text1.hwnd, CustomDraw:=mnuTest(3).Checked)
            mnuTest(0).Enabled = False
            mnuTest(1).Enabled = True
        
        Case 1 '-- Uninitialize
            Call m_oSB.UninitializeCoolSB
            mnuTest(0).Enabled = True
            mnuTest(1).Enabled = False
        
        Case 3 '-- Enable/disable custom draw
            mnuTest(3).Checked = Not mnuTest(3).Checked
            If (m_oSB.Initialized) Then
                m_oSB.CustomDraw = mnuTest(3).Checked
            End If
    End Select
End Sub

Private Sub Form_Resize()
    Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

'//

Private Sub m_oSB_OnPaint( _
            ByVal hDC As Long, _
            ByVal ScrollBarID As Long, _
            ByVal x1 As Long, ByVal y1 As Long, _
            ByVal x2 As Long, ByVal y2 As Long, _
            ByVal Part As csbOnPaintPartCts, _
            ByVal Pressed As Boolean _
            )

  Dim lfArrowDir As Long
    
    '-- Here, you can paint whatever you want...
    '   I'm painting a simple colored soft-3D scrollbar
    
    Select Case Part
        
        Case [ppTLButton], [ppBRButton]
            
            lfArrowDir = [afDFCS_SCROLLUP] + -(Part = [ppBRButton]) + 2 * -(ScrollBarID = 0)
                
            If (m_oSB.Enabled(ScrollBarID)) Then
                If (Pressed) Then
                    Call DrawArrowEx(hDC, x1, y1, x2, y2, lfArrowDir, ShiftColor(clrMYCOLOR, -50), ShiftColor(clrMYCOLOR, 75))
                    Call DrawRectEx(hDC, x1, y1, x2, y2, ShiftColor(clrMYCOLOR, -75))
                  Else
                    Call DrawArrowEx(hDC, x1, y1, x2, y2, lfArrowDir, clrMYCOLOR, ShiftColor(clrMYCOLOR, -75))
                    Call DrawEdgeEx(hDC, x1, y1, x2, y2, clrMYCOLOR)
                End If
              Else
                Call DrawArrowEx(hDC, x1, y1, x2, y2, lfArrowDir, TranslateColor(vbButtonFace), TranslateColor(vb3DShadow))
                Call DrawEdgeEx(hDC, x1, y1, x2, y2, TranslateColor(vbButtonFace))
            End If
            
        Case [ppTLTrack], [ppBRTrack]
        
            If (m_oSB.Enabled(ScrollBarID)) Then
                If (Pressed) Then
                    Call FillRectEx(hDC, x1, y1, x2, y2, ShiftColor(clrMYCOLOR, -50))
                  Else
                    Call FillRectEx(hDC, x1, y1, x2, y2, ShiftColor(clrMYCOLOR, 50))
                End If
              Else
                Call FillRectEx(hDC, x1, y1, x2, y2, TranslateColor(vb3DFace))
            End If
            
        Case [ppNullTrack]
        
            If (m_oSB.Enabled(ScrollBarID)) Then
                Call FillRectEx(hDC, x1, y1, x2, y2, ShiftColor(clrMYCOLOR, 25))
              Else
                Call FillRectEx(hDC, x1, y1, x2, y2, TranslateColor(vb3DFace))
            End If
            
        Case [ppThumb]
        
            If (m_oSB.Enabled(ScrollBarID)) Then
                If (Pressed) Then
                    Call DrawRectEx(hDC, x1, y1, x2, y2, ShiftColor(clrMYCOLOR, -75))
                    Call FillRectEx(hDC, x1 + 1, y1 + 1, x2 - 1, y2 - 1, ShiftColor(clrMYCOLOR, -50))
                  Else
                    Call DrawEdgeEx(hDC, x1, y1, x2, y2, clrMYCOLOR)
                    Call FillRectEx(hDC, x1 + 1, y1 + 1, x2 - 1, y2 - 1, clrMYCOLOR)
                End If
              Else
                Call DrawRectEx(hDC, x1, y1, x2, y2, TranslateColor(vb3DShadow))
                Call FillRectEx(hDC, x1 + 1, y1 + 1, x2 - 1, y2 - 1, TranslateColor(vbButtonFace))
            End If
            
        Case [ppSizer]

            If (m_oSB.HasSizer) Then
                Call DrawArrowEx(hDC, x1, y1, x2, y2, [afDFCS_SCROLLSIZEGRIP], clrMYCOLOR, ShiftColor(clrMYCOLOR, -50))
              Else
                Call FillRectEx(hDC, x1, y1, x2, y2, clrMYCOLOR)
            End If
    End Select
End Sub
