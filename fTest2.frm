VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fTest2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo 2"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   345
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   609
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "ImageCombo1"
      ImageList       =   "ImageList1"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fTest2.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   240
      Left            =   480
      TabIndex        =   4
      Top             =   975
      Value           =   1  'Checked
      Width           =   1140
   End
   Begin VB.ComboBox Combo3 
      Height          =   330
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   1905
   End
   Begin VB.ComboBox Combo2 
      Height          =   330
      Left            =   2520
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   1560
      Width           =   1905
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1905
   End
End
Attribute VB_Name = "fTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oFlatCombo1      As New cFlatCombo
Private m_oFlatCombo2      As New cFlatCombo
Private m_oFlatCombo3      As New cFlatCombo
Private m_oFlatImageCombo1 As New cFlatImageCombo

Private Sub Form_Load()

  Dim i As Long
    
    For i = 1 To 50
        Combo1.AddItem "Item " & i
        Combo2.AddItem "Item " & i
        Combo3.AddItem "Item " & i
        ImageCombo1.ComboItems.Add , , "Item " & i, 1, 2
    Next i

    m_oFlatCombo1.Initialize Combo1
    m_oFlatCombo2.Initialize Combo2
    m_oFlatCombo3.Initialize Combo3
    m_oFlatImageCombo1.Initialize ImageCombo1
End Sub

Private Sub chkEnabled_Click()

    Combo1.Enabled = -chkEnabled
    Combo2.Enabled = -chkEnabled
    Combo3.Enabled = -chkEnabled
    ImageCombo1.Enabled = -chkEnabled
End Sub
