VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   5175
   ClientTop       =   3060
   ClientWidth     =   3180
   DrawMode        =   7  'Invert
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmtest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Dialog Position"
      Height          =   1950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3195
      Begin ColorBox_.ColorBox ColorBox1 
         Left            =   2250
         Top             =   630
         _ExtentX        =   847
         _ExtentY        =   503
         AdjustPosition  =   0   'False
         DialogStartUpPosition=   2
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Adjust Position."
         Height          =   255
         Left            =   330
         TabIndex        =   8
         Top             =   1605
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2460
         TabIndex        =   5
         Text            =   "-10"
         Top             =   1440
         Width           =   585
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   330
         Left            =   2460
         TabIndex        =   4
         Text            =   "800"
         Top             =   1080
         Width           =   585
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Manual"
         Height          =   285
         Left            =   300
         TabIndex        =   3
         Top             =   1245
         Width           =   1125
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Centre Screen"
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   900
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Centre Owner"
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   555
         Width           =   2265
      End
      Begin VB.Label lblC 
         BackColor       =   &H00F0B7A8&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   4
         Left            =   2475
         TabIndex        =   13
         Top             =   90
         Width           =   465
      End
      Begin VB.Label lblC 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   3
         Left            =   1935
         TabIndex        =   12
         Top             =   90
         Width           =   465
      End
      Begin VB.Label lblC 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   2
         Left            =   1395
         TabIndex        =   11
         Top             =   90
         Width           =   465
      End
      Begin VB.Label lblC 
         BackColor       =   &H00BEE4CA&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   855
         TabIndex        =   10
         Top             =   90
         Width           =   465
      End
      Begin VB.Label lblC 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   90
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Y :"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2160
         TabIndex        =   7
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "X :"
         Enabled         =   0   'False
         Height          =   225
         Left            =   2160
         TabIndex        =   6
         Top             =   1140
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Dim BkColor As Long
Dim FkColor As Long
Dim SelIndex As Integer
Private Declare Function OleTranslateColor Lib "olepro32" (clr As OLE_COLOR, hpal As Long, pcolorref As Long) As Long

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        ColorBox1.AdjustPosition = True
    Else
        ColorBox1.AdjustPosition = False
    End If
End Sub



Private Sub ColorBox1_ColorChange(ByVal NewColor As Long)
        lblC(SelIndex).BackColor = NewColor
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
            Unload Form1
    End If
     
End Sub

Private Sub Form_Load()
    BkColor = vbWhite
    Me.FontSize = 45
    Me.FontBold = True
    Text1.Text = (Screen.Width / 15) - 250
End Sub




Private Sub lblC1_Click()
    
End Sub

 

Private Sub lblC_Click(Index As Integer)
    SelIndex = Index
    With ColorBox1
    Select Case True
    Case Option3.Value
        .DialogStartUpPosition = CentreOwner
    Case Option4.Value
        .DialogStartUpPosition = CentreScreen
    Case Option5.Value
        .DialogStartUpPosition = Manual
    End Select
    .DialogLeft = Val(Text1.Text)
    .DialogTop = Val(Text2.Text)
    .Color = lblC(Index).BackColor
    .Show
    End With
    lblC(Index).BackColor = ColorBox1.Color

End Sub

Private Sub Option3_Click()
    ColorBox1.DialogStartUpPosition = CentreOwner
    Text1.Enabled = False
    Text2.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False

End Sub

Private Sub Option4_Click()
    ColorBox1.DialogStartUpPosition = CentreScreen
    Text1.Enabled = False
    Text2.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False

End Sub

Private Sub Option5_Click()
    ColorBox1.DialogStartUpPosition = Manual
    Text1.Enabled = True
    Text2.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
End Sub

