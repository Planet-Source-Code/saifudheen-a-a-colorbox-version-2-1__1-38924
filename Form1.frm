VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ColorBox"
   ClientHeight    =   4140
   ClientLeft      =   3525
   ClientTop       =   2730
   ClientWidth     =   6510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0CCA
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "O&K"
      Height          =   315
      Left            =   5550
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   2760
      Width           =   825
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Preset"
      Height          =   315
      Left            =   5550
      TabIndex        =   7
      Top             =   3180
      Width           =   825
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   " << Add "
      Height          =   315
      Left            =   4560
      TabIndex        =   23
      Top             =   3600
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ComboBox cmbPreset 
      Height          =   315
      Left            =   4650
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   4590
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   -30
      Width           =   885
      Begin VB.Shape Shape1 
         Height          =   810
         Left            =   45
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   45
         TabIndex        =   11
         Top             =   540
         Width           =   780
      End
      Begin VB.Label lblSelColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   45
         TabIndex        =   10
         Top             =   135
         Width           =   780
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   5550
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   3600
      Width           =   825
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   4590
      TabIndex        =   12
      Top             =   960
      Width           =   2055
      Begin VB.CheckBox chkWeb 
         Caption         =   "Only web colors."
         Height          =   195
         Left            =   30
         TabIndex        =   34
         Top             =   135
         Width           =   1770
      End
      Begin VB.TextBox txtK 
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   29
         Text            =   "100"
         ToolTipText     =   "K"
         Top             =   2745
         Width           =   405
      End
      Begin VB.TextBox txtYellow 
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   28
         Text            =   "100"
         ToolTipText     =   "Yellow"
         Top             =   2385
         Width           =   405
      End
      Begin VB.TextBox txtMagenta 
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   27
         Text            =   "100"
         ToolTipText     =   "Magenta"
         Top             =   2025
         Width           =   405
      End
      Begin VB.TextBox txtCyan 
         Height          =   285
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   26
         Text            =   "0"
         ToolTipText     =   "Cyan"
         Top             =   1665
         Width           =   405
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "0"
         ToolTipText     =   "Blue"
         Top             =   1200
         Width           =   405
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "Green"
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "255"
         ToolTipText     =   "Red"
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   450
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Hue"
         Top             =   480
         Width           =   405
      End
      Begin VB.TextBox txtS 
         Height          =   285
         Left            =   450
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   4
         Text            =   "100"
         ToolTipText     =   "Saturation"
         Top             =   840
         Width           =   405
      End
      Begin VB.TextBox txtB 
         Height          =   285
         Left            =   450
         MaxLength       =   3
         MousePointer    =   4  'Icon
         TabIndex        =   5
         Text            =   "100"
         ToolTipText     =   "Brightness"
         Top             =   1200
         Width           =   405
      End
      Begin VB.OptionButton optH 
         Caption         =   "H:"
         Height          =   255
         Left            =   0
         TabIndex        =   0
         ToolTipText     =   "Hue"
         Top             =   540
         Width           =   465
      End
      Begin VB.OptionButton optS 
         Caption         =   "S:"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Saturation"
         Top             =   885
         Width           =   465
      End
      Begin VB.OptionButton optB 
         Caption         =   "B:"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Brightness"
         Top             =   1230
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "K:"
         Height          =   195
         Left            =   45
         TabIndex        =   33
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Y:"
         Height          =   195
         Left            =   45
         TabIndex        =   32
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "M:"
         Height          =   195
         Left            =   45
         TabIndex        =   31
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "C:"
         Height          =   240
         Left            =   45
         TabIndex        =   30
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label lblR 
         Caption         =   "R:"
         Height          =   225
         Left            =   1140
         TabIndex        =   21
         ToolTipText     =   "Red"
         Top             =   540
         Width           =   195
      End
      Begin VB.Label lblG 
         Caption         =   "G:"
         Height          =   225
         Left            =   1140
         TabIndex        =   20
         ToolTipText     =   "Green"
         Top             =   885
         Width           =   225
      End
      Begin VB.Label lblB 
         Caption         =   "B:"
         Height          =   225
         Left            =   1140
         TabIndex        =   19
         ToolTipText     =   "Blue"
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label lblS 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   18
         Top             =   870
         Width           =   165
      End
      Begin VB.Label lblBB 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   900
         TabIndex        =   17
         Top             =   1230
         Width           =   195
      End
      Begin VB.Label lblH 
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   900
         TabIndex        =   16
         Top             =   510
         Width           =   105
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   4530
      TabIndex        =   24
      Top             =   3060
      Width           =   975
      Begin VB.Label lblADDColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   30
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   825
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'************************ColorBox OCX Version 2.1************************
'Main dialog
'Author : Saifudheen A.A.
'E-mail : keraleeyan@msn.com.
' Suggestions, Votes all are welcome.
'********************************************************************
Enum sMode
    Picker = 0
    About = 1
    Custom = 2
    SafeColor = 3
End Enum

Dim MainBoxHit  As Boolean
Dim SelectBoxHit As Boolean
Public Mode As sMode
Public UserMode As Integer 'usercontrol Mode variable
Dim HueEntering As Boolean
Dim SaturationEntering As Boolean
Dim BrightnessEntering As Boolean
Dim OldHue As Integer, OldSaturation As Integer, OldBrightness As Integer
Dim ChangeByInput As Boolean
Dim Loaded As Boolean
Dim LcButtonRect As RECT
Dim LcSmpRect As RECT
Dim lcButtonPressed As Boolean

Public Sub LoadMode(ByVal lMode As Integer)
    Dim cl As Long
    Dim mc As Long
    Dim hs As HSB, hs1 As HSB
    Dim red As Integer, green As Integer, blue As Integer
    pMode = lMode
    cl = lblSelColor.BackColor ' m_Color
    GetRGB cl, red, green, blue
    Form1.Cls
    PrintLastColor
    Select Case lMode
    Case 0
        LoadMainHue
        hs = RGBtoHSB(lblSelColor.BackColor)
        SelectedMainPos = MainBox.Top + (360 - hs.Hue) * 255 / 360
        Call DrawSlider(SelectedMainPos)
        mc = GetPixel(Me.hdc, MainBox.Left + 5, SelectedMainPos)
        GetRGB mc, red, green, blue
        SelectedPos.X = SelectBox.Left + hs.Saturation * 255 / 100
        SelectedPos.Y = SelectBox.Bottom - hs.Brightness * 255 / 100
        
        LoadVariantsHue red, green, blue
        DrawPicker
    Case 1
        hs = RGBtoHSB(lblSelColor.BackColor)
        hs1 = hs
        LoadVariantsSaturation hs.Saturation / 100
        SelectedPos.X = SelectBox.Left + hs.Hue * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - (hs.Brightness) * 255 / 100
        SelectedMainPos = MainBox.Top + (100 - hs.Saturation) * 255 / 100
        Call DrawSlider(SelectedMainPos)
        hs1.Saturation = 100
        hs1.Brightness = 100
        HSBtoRGB hs1, red, green, blue
        DrawPicker
        LoadMainSaturation Me.hdc, red, green, blue, hs.Brightness / 100
    Case 2
        hs = RGBtoHSB(lblSelColor.BackColor)
        hs1 = hs
        LoadVariantsBrightness hs.Brightness / 100
        SelectedPos.X = SelectBox.Left + hs.Hue * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - hs.Saturation * 255 / 100
        SelectedMainPos = MainBox.Top + 255 - (hs.Brightness * 255 / 100)
        Call DrawSlider(SelectedMainPos)
        'hs1.Saturation = 100
        hs1.Brightness = 100
        HSBtoRGB hs1, red, green, blue
        LoadMainBrightness Me.hdc, red, green, blue
        DrawPicker
    End Select
    mc = lblSelColor.BackColor  'm_Color
    GetRGB mc, red, green, blue
    Text1.Text = red
    Text2.Text = green
    Text3.Text = blue
    txtH.Text = CInt(hs.Hue)
    txtS.Text = CInt(hs.Saturation)
    txtB.Text = CInt(hs.Brightness)
    Me.PSet (-100, -100)
End Sub

Private Sub chkWeb_Click()
    m_WebColors = IIf(chkWeb.Value = vbChecked, True, False)
    If Loaded = False Then Exit Sub
    ReloadColors
End Sub

Private Sub ReloadColors()
    Dim cl As Long
    Dim red As Integer, green As Integer, blue As Integer
    DrawPicker
    DrawSlider SelectedMainPos
    Dim hs As HSB
    Select Case True
    Case optH.Value
        SelectedMainPos = MainBox.Top + (360 - Val(txtH.Text)) * 255 / 360
        SelectedPos.X = SelectBox.Left + Val(txtS.Text) * 255 / 100
        SelectedPos.Y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100
        LoadMainHue
        cl = GetPixel(Form1.hdc, MainBox.Left + 5, SelectedMainPos)
        GetRGB cl, red, green, blue
        LoadVariantsHue red, green, blue
        DrawPicker
    Case optS.Value
        SelectedMainPos = MainBox.Top + (100 - Val(txtS.Text)) * 255 / 100
        SelectedPos.X = SelectBox.Left + Val(txtH.Text) * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - Val(txtB.Text) * 255 / 100
        LoadVariantsSaturation Val(txtS.Text) / 100
        hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
        HSBtoRGB hs, red, green, blue
        LoadMainSaturation Form1.hdc, red, green, blue, Val(txtB.Text) / 100
        DrawPicker
    Case optB.Value
        SelectedMainPos = MainBox.Top + (100 - Val(txtB.Text)) * 255 / 100
        SelectedPos.X = SelectBox.Left + Val(txtH.Text) * 255 / 360
        SelectedPos.Y = SelectBox.Bottom - Val(txtS.Text) * 255 / 100
        LoadVariantsBrightness Val(txtB.Text) / 100
        hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
        HSBtoRGB hs, red, green, blue
        LoadMainBrightness Form1.hdc, red, green, blue
        DrawPicker
    End Select
    DrawSlider SelectedMainPos

End Sub
Private Sub cmbPreset_Click()
Form1.Cls
Form1.PrintLastColor
Dim r As Integer, g As Integer, b As Integer
Dim cl As Long
Dim HexV As String
On Error GoTo r:
Select Case cmbPreset.ListIndex
Case 0
    lblADDColor.Visible = False
    cmdADD.Visible = False
    Mode = About
    PrintAbout Form1.hdc
    
Case 1
    pMode = 3
    Me.DrawStyle = 0
    LoadCustomColors
    cmdADD.Visible = True
    lblADDColor.Visible = True
    Mode = Custom
    cl = lblSelColor.BackColor
    GetRGB cl, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV
Case 2
    pMode = 4
    Me.DrawStyle = 0
    LoadSafePalette
    cmdADD.Visible = False
    lblADDColor.Visible = False
    Mode = SafeColor
    cl = lblSelColor.BackColor
    GetRGB cl, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV
End Select
Me.PSet (-100, -100)
Exit Sub
r:
Exit Sub
'MsgBox Error
End Sub

Private Sub cmdADD_Click()
    svdColor(cPaletteIndex) = lblADDColor.BackColor
    DrawSafePicker cPaletteIndex, True 'Erase
    Form1.DrawWidth = 1
    Form1.DrawMode = 13
    Form1.FillStyle = 0
    Form1.ForeColor = 0
    Form1.FillColor = svdColor(cPaletteIndex)
    Rectangle Form1.hdc, Preset(cPaletteIndex).Left, Preset(cPaletteIndex).Top, Preset(cPaletteIndex).Right, Preset(cPaletteIndex).Bottom
    SaveCustomColors
    DrawSafePicker cPaletteIndex, False
    lblSelColor.BackColor = svdColor(cPaletteIndex)
    Form1.PSet (-100, -100)
End Sub

Private Sub Command1_Click()
    If Mode = About Then cmbPreset.ListIndex = 2
    m_Color = lblSelColor.BackColor
    LastColor = m_Color
    Me.Hide
End Sub



Private Sub Command2_Click()
    Me.lblSelColor.BackColor = mOldColor
    Form1.Hide
End Sub

Sub Command3_Click()
Me.Cls
Dim r As Integer, g As Integer, b As Integer
Dim HexV As String
optNotClicked = False
If Mode = Picker Then
    Frame2.Visible = False
    Mode = Custom
    Command3.Caption = "&Picker"
    cmbPreset.Visible = True
    LoadCustomColors
    cmbPreset.ListIndex = 1
    lblADDColor.BackColor = lblSelColor.BackColor
    cmdADD.Visible = True
    lblADDColor.Visible = True
    GetRGB lblSelColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV
Else
    'Me.Cls
    cmbPreset.Visible = False
    Command3.Caption = "&Preset"
    'DrawPicker
     
    Select Case True
    Case optH
         optH_Click
    Case optS
        optS_Click
    Case optB
        optB_Click
    End Select
    Frame2.Visible = True
    cmdADD.Visible = False
    lblADDColor.Visible = False

    Mode = Picker
End If


End Sub



Function GetSafeColor(Index As Integer, r As Integer, g As Integer, b As Integer, Optional HexVal As String) As Long
Dim i As Long, j As Long, k As Long
Dim count As Integer
Dim strR As String, strG As String, strB As String

For i = 0 To &HFF Step &H33
    For j = 0 To &HFF Step &H33
        For k = 0 To &HFF Step &H33
            count = count + 1
            If count = Index Then
                r = i: g = j: b = k
                GetSafeColor = RGB(i, j, k)
                GetHexVal r, g, b, HexVal
                Exit Function
            End If
           
        Next k
    Next j
Next i

End Function

Sub GetHexVal(red As Integer, green As Integer, blue As Integer, strHex As String)
    Dim strR As String, strG As String, strB As String
    strR = Trim(Hex(red))
        If Len(strR) = 1 Then strR = "0" & strR
    strG = Trim(Hex(green))
        If Len(strG) = 1 Then strR = "0" & strR
    strB = Trim(Hex(blue))
        If Len(strB) = 1 Then strR = "0" & strR
        strHex = strR & strG & strB
End Sub


Public Sub PrintLastColor(Optional ByVal Pressed As Boolean)
    Me.DrawWidth = 1
    Me.DrawMode = 13
    Me.FillStyle = 0
    Me.FillColor = Form1.BackColor
    Me.ForeColor = Form1.BackColor
    Rectangle Me.hdc, LcButtonRect.Left, LcButtonRect.Top, LcButtonRect.Right, LcButtonRect.Bottom
    
    Me.CurrentX = LcSmpRect.Right + 9
    Me.CurrentY = LcSmpRect.Top + 0
    If Pressed Then
        Me.CurrentX = Me.CurrentX + 1
        Me.CurrentY = Me.CurrentY + 1
    End If
    Me.ForeColor = 0
    Me.Print "OK"
    Me.CurrentY = LcButtonRect.Top - 15
    Me.CurrentX = LcButtonRect.Left + 2
    Me.Print "Last color"
    Me.FillStyle = 0
    Me.FillColor = LastColor
    Me.ForeColor = 0
    If Pressed Then
        Rectangle Me.hdc, LcSmpRect.Left + 1, LcSmpRect.Top + 1, LcSmpRect.Right + 1, LcSmpRect.Bottom + 1
        DrawEdge Me.hdc, LcButtonRect, BDR_SUNKENOUTER Or BDR_SUNKENINNER, BF_RECT
    Else
        Rectangle Me.hdc, LcSmpRect.Left, LcSmpRect.Top, LcSmpRect.Right, LcSmpRect.Bottom
        DrawEdge Me.hdc, LcButtonRect, BDR_RAISEDOUTER Or BDR_RAISEDINNER, BF_RECT
    End If
    
    Me.PSet (-100, -100)
End Sub




Private Sub Form_Load()
    
    ' Set these Parameters on Basis of where  should Select Box and  Main Box Should appear
    Dim OldP As POINTAPI
    ReDim Preset(1 To 224)  'RECT structure
    Preset(1).Left = 10
    Preset(1).Top = 10
    Preset(1).Right = 25
    Preset(1).Bottom = 25
    
    With LcButtonRect
        .Right = Me.ScaleWidth - 8
        .Left = .Right - 55
        .Top = 25
        .Bottom = .Top + 20
    End With
    
    With LcSmpRect
        .Left = LcButtonRect.Left + 4
        .Top = LcButtonRect.Top + 4
        .Bottom = LcButtonRect.Bottom - 4
        .Right = .Left + .Bottom - .Top + 6
    End With
    
    Mode = Picker   'Initialize Mode as  normal Picker
    With lpBI.bmiHeader
        .biBitCount = 24
        .biCompression = 0&
        .biPlanes = 1
        .biSize = Len(lpBI.bmiHeader)
    End With
    
    '// Setting the position of RECTS for Safe Colors
    For i = 2 To 224
        If i Mod 16 = 0 Then
            Preset(i).Top = Preset(i - 1).Top
            Preset(i).Left = Preset(i - 1).Right + 3
            Preset(i).Bottom = Preset(i).Top + 15
            Preset(i).Right = Preset(i).Left + 15
            If i = 224 Then GoTo Jump
            i = i + 1
            Preset(i).Top = Preset(i - 1).Bottom + 3
            Preset(i).Left = Preset(1).Left
        Else
            Preset(i).Top = Preset(i - 1).Top
            Preset(i).Left = Preset(i - 1).Right + 3
        End If
Jump:
        Preset(i).Bottom = Preset(i).Top + 15
        Preset(i).Right = Preset(i).Left + 15
    Next i

    
    k = 255
    'Show
    SelectBox.Left = 10
    SelectBox.Top = 10
    SelectBox.Right = SelectBox.Left + k
    SelectBox.Bottom = SelectBox.Top + k
    MainBox.Left = SelectBox.Right + 12
    MainBox.Top = SelectBox.Top
    MainBox.Right = MainBox.Left + 15
    MainBox.Bottom = SelectBox.Bottom
     
    With cmbPreset
    .AddItem "About ColorBox"
    .AddItem "Custom..."
    .AddItem "Safe Palette (216)"
    End With
    
    'LoadMode pMode
    
    cPaletteIndex = 1

    Me.ForeColor = vbBlack
    If m_WebColors Then
        Form1.chkWeb.Value = vbChecked
    Else
        Form1.chkWeb.Value = vbUnchecked
    End If
    Loaded = True
End Sub




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If X >= LcSmpRect.Left And X <= LcSmpRect.Right And Y >= LcSmpRect.Top And Y <= LcSmpRect.Bottom Then
            If pMode < 3 Then
                lblSelColor.BackColor = LastColor
                LoadMode pMode
            End If
            Exit Sub
        End If
        If X >= LcButtonRect.Left And X <= LcButtonRect.Right And Y >= LcButtonRect.Top And Y <= LcButtonRect.Bottom Then
            PrintLastColor True
            Me.PSet (-100, -100)
            lcButtonPressed = True
            Exit Sub
        End If
        
        If Mode = Picker Then
            If X >= SelectBox.Left And X <= SelectBox.Right And Y >= SelectBox.Top And Y <= SelectBox.Bottom Then
                '// In SelectBox Boundary
                SelectBoxHit = True
                Me.MousePointer = vbCustom
               Call MouseOnSelectBox(X, Y)
            End If
            If X >= MainBox.Left And X < MainBox.Right + 11 And Y >= MainBox.Top - 2 And Y < MainBox.Bottom + 3 Then
                '// In MainBox Boundary
                MainBoxHit = True
                If Y > MainBox.Bottom Then Y = MainBox.Bottom
                If Y < MainBox.Top Then Y = MainBox.Top
                Call MouseOnMainBox(X, Y)
            End If
        Else
            If Mode <> About Then
                HandlePresetValues X, Y, Mode
            End If
        End If
    End If
End Sub

Sub DrawSafePicker(Index As Integer, Clear As Boolean)
    Dim r As RECT
    Dim L As Long
    Me.FillStyle = 1
    Me.DrawMode = 13
    Me.DrawWidth = 3
    If Clear Then
        Me.ForeColor = Me.BackColor
        Rectangle Me.hdc, Preset(Index).Left - 2, Preset(Index).Top - 2, Preset(Index).Right + 2, Preset(Index).Bottom + 2
    Else
        r.Left = Preset(Index).Left - 3
        r.Top = Preset(Index).Top - 3
        r.Right = Preset(Index).Right + 3
        r.Bottom = Preset(Index).Bottom + 3
        Call DrawEdge(Form1.hdc, r, BDR_SUNKENINNER Or BDR_SUNKENOUTER, BF_RECT Or BF_SOFT)
    End If
    
End Sub

Private Sub MouseOnSelectBox(X As Single, Y As Single)
            Dim cl As Long
            Dim hs As HSB
            Dim r As Integer
            Dim g As Integer
            Dim b As Integer
            DrawPicker
            SelectedPos.X = X
            SelectedPos.Y = Y
            DrawPicker
            cl = GetPixel(Me.hdc, X, Y)
            GetRGB cl, r, g, b
            If optS.Value Then
                txtB.Text = Int((255 - SelectedPos.Y + SelectBox.Top) * 100 / 255) 'Brightness Level
                txtH.Text = Int((SelectedPos.X - SelectBox.Left) * 360 / 255) 'Hue Level
                hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
                HSBtoRGB hs, r, g, b
                LoadMainSaturation Form1.hdc, r, g, b, Val(txtB.Text) / 100
            End If
            
            If optB.Value Then
                hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
                HSBtoRGB hs, r, g, b
                LoadMainBrightness Form1.hdc, r, g, b
                txtS.Text = Int((SelectBox.Bottom - SelectedPos.Y) * 100 / 255)    ' Saturation Level
                txtH.Text = Int((SelectedPos.X - SelectBox.Left) * 360 / 255)   'Hue Level
            End If

            If optH.Value Then
                txtS.Text = Int((SelectedPos.X - SelectBox.Left) * 100 / 255)   ' Saturation Level
                txtB.Text = Int((255 - SelectedPos.Y + SelectBox.Top) * 100 / 255) 'Brightness Level
            Else
                cl = GetPixel(Me.hdc, MainBox.Left + 5, SelectedMainPos)
                GetRGB cl, r, g, b
            End If
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
            lblSelColor.BackColor = cl

End Sub

Private Sub MouseOnMainBox(X As Single, Y As Single)
        Dim cl As Long
        Dim r As Integer
        Dim g As Integer
        Dim b As Integer

            DrawSlider SelectedMainPos
            DrawSlider Y
            SelectedMainPos = Y
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            GetRGB cl, r, g, b
            
            If optH.Value Then
                txtH.Text = Int((255 - Y + SelectBox.Top) * 360 / 255)
                GetColorFromHSB Val(txtS.Text), Val(txtB.Text)
            Else
                Text1.Text = r
                Text2.Text = g
                Text3.Text = b
                lblSelColor.BackColor = cl
            End If
            If optS.Value Then
                txtS.Text = Int((255 - Y + SelectBox.Top) * 100 / 255)
                
            End If
            If optB.Value Then
                txtB.Text = Int((255 - Y + SelectBox.Top) * 100 / 255)
            End If

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt As POINTAPI
    Dim ClipRect As RECT
    pt.X = X
    pt.Y = Y
    
    If Button = 1 Then
        If lcButtonPressed Then
            If X < LcButtonRect.Left Or X > LcButtonRect.Right Or Y < LcButtonRect.Top Or Y > LcButtonRect.Bottom Then
                PrintLastColor False
                Me.PSet (-100, -100)
                lcButtonPressed = False
                Exit Sub
            End If
        End If

        If Mode = Picker Then
            If SelectBoxHit Then
                
                If X < SelectBox.Left Then X = SelectBox.Left
                If X > SelectBox.Right Then X = SelectBox.Right
                If Y < SelectBox.Top Then Y = SelectBox.Top
                If Y > SelectBox.Bottom Then Y = SelectBox.Bottom
                '// In SelectBox Region
                
                pt.X = 10
                pt.Y = 10
                ClientToScreen Me.hwnd, pt
                ClipRect.Left = pt.X
                ClipRect.Top = pt.Y
                
                pt.X = 266
                pt.Y = 266
                ClientToScreen Me.hwnd, pt
                ClipRect.Right = pt.X
                ClipRect.Bottom = pt.Y
                ClipCursor ClipRect
                Call MouseOnSelectBox(X, Y)
            End If
            
            If MainBoxHit Then
                '// In MainBox region
                X = MainBox.Left + 2
                If Y > MainBox.Bottom Then Y = MainBox.Bottom
                If Y < MainBox.Top Then Y = MainBox.Top
                Call MouseOnMainBox(X, Y)
            End If
        Else
            If Mode <> About Then
                HandlePresetValues X, Y, Mode
            End If
        End If
    End If
End Sub

Private Sub HandlePresetValues(X As Single, Y As Single, pMode As sMode)
            Dim i As Integer
            Dim r As Integer, g As Integer, b As Integer
            Dim HexV As String
            For i = 1 To 224
                If X > Preset(i).Left And X < Preset(i).Right And Y > Preset(i).Top And Y < Preset(i).Bottom Then
                    DrawSafePicker cPaletteIndex, True
                    DrawSafePicker i, False
                    Me.PSet (-100, -100)
                    cPaletteIndex = i
                    Select Case pMode
                    Case 1
                        
                    Case 2
                        lblSelColor.BackColor = svdColor(i)
                        GetRGB svdColor(i), r, g, b
                        GetHexVal r, g, b, HexV
                    Case 3
                        lblSelColor.BackColor = GetSafeColor(i, r, g, b, HexV)
                        
                    End Select
                    
                    PrintRGBHEX r, g, b, HexV
                    Exit Sub
                End If
            Next i

End Sub

Private Sub PrintRGBHEX(r As Integer, g As Integer, b As Integer, HexV As String)
    Me.FontSize = 8
    Me.FillColor = Me.BackColor
    Me.ForeColor = Me.BackColor
    Me.DrawMode = 13
    Me.Line (MainBox.Right + 20, 90)-(MainBox.Right + 20 + 190, 90 + 50), Me.BackColor, BF
    Me.CurrentX = MainBox.Right + 20
    Me.CurrentY = 90
    Me.ForeColor = 0
    Me.Print "R: " & r
    Me.CurrentX = MainBox.Right + 20
    Me.Print "G: " & g
    Me.CurrentX = MainBox.Right + 20
    Me.Print "B: " & b
    Me.CurrentX = MainBox.Right + 20
    Me.Print "Hex: #" & HexV

End Sub

Private Sub GetColorFromHSB(ByVal Sat As Integer, ByVal br As Integer)
    '//
    '// This Function Evaluates the Resulting Color Value While Sliding the Hue Shades From Insisting Brightness and Saturation Values
    Dim r As Integer, g As Integer, b As Integer
    Dim red As Integer, green As Integer, blue As Integer
    Dim cl As Long
    Dim X As Long, Y As Long
    On Error Resume Next
    X = Sat * (255 / 100)
    cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
    GetRGB cl, red, green, blue
    r = (red + ((255 - red) * (100 - Sat)) / 100) * br / 100
    g = (green + ((255 - green) * (100 - Sat)) / 100) * br / 100
    b = (blue + ((255 - blue) * (100 - Sat)) / 100) * br / 100
    lblSelColor.BackColor = RGB(r, g, b)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim em As RECT
    em.Right = Screen.Width / 15
    em.Bottom = Screen.Height / 15
    ClipCursor em
    SelectBoxHit = False
    Me.MousePointer = vbDefault
    PrintLastColor False
    
    If X >= LcSmpRect.Left And X <= LcSmpRect.Right And Y >= LcSmpRect.Top And Y <= LcSmpRect.Bottom Then
         Exit Sub
    End If
    
    If lcButtonPressed And X >= LcButtonRect.Left And X <= LcButtonRect.Right And Y >= LcButtonRect.Top And Y <= LcButtonRect.Bottom Then
         lblSelColor.BackColor = LastColor
         Command1_Click
         lcButtonPressed = False
    End If
    If MainBoxHit And optH.Value Then
        Dim r As Integer, g As Integer, b As Integer
        Dim cl As Long
        Dim px As Long
        DrawPicker
        cl = GetPixel(Me.hdc, MainBox.Left + 3, SelectedMainPos)
        GetRGB cl, r, g, b
        LoadVariantsHue r, g, b
        DrawPicker
        cl = Me.Point(SelectedPos.X, SelectedPos.Y)
        px = cl
        GetRGB cl, r, g, b
        Text1.Text = r
        Text2.Text = g
        Text3.Text = b
        lblSelColor.BackColor = Me.Point(SelectedPos.X, SelectedPos.Y)
    End If
    If MainBoxHit And optS.Value Then
        DrawPicker
        LoadVariantsSaturation Val(txtS.Text) / 100
        DrawPicker
    End If
    If MainBoxHit And optB.Value Then
        DrawPicker
        LoadVariantsBrightness Val(txtB.Text) / 100
        DrawPicker
    End If
    Me.PSet (-100, -100)

    MainBoxHit = False
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        If UnloadMode = vbFormControlMenu Then
            Cancel = True
            Me.Hide
        End If
End Sub

Private Sub Label2_Click()
    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    lblSelColor.BackColor = Label2.BackColor
    If pMode < 3 Then
        LoadMode pMode
    Else
        GetRGB Label2.BackColor, r, g, b
        GetHexVal r, g, b, HexV
        PrintRGBHEX r, g, b, HexV
    End If
End Sub

Private Sub Label2_DblClick()
    lblSelColor.BackColor = Label2.BackColor

End Sub

Private Sub lblADDColor_Click()
    On Error Resume Next

    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    GetRGB lblADDColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV
End Sub

Private Sub lblSelColor_Click()
    On Error Resume Next

    Dim r As Integer, g As Integer, b As Integer
    Dim HexV As String
    GetRGB lblSelColor.BackColor, r, g, b
    GetHexVal r, g, b, HexV
    PrintRGBHEX r, g, b, HexV

End Sub

Public Sub optB_Click()
    On Error Resume Next
    If optNotClicked Then Exit Sub
    pMode = 2
    LoadMode pMode
End Sub

Public Sub optH_Click()
    On Error Resume Next
    pMode = 0
    If optNotClicked Then Exit Sub
    LoadMode pMode
End Sub

Public Sub optS_Click()
    If optNotClicked Then Exit Sub
    pMode = 1
    LoadMode pMode
End Sub


 
Sub PrintAbout(hdc As Long)
    On Error Resume Next

    Dim qrc As RECT, brdr As RECT
    qrc.Left = 20
    qrc.Top = 90
    qrc.Right = 280
    qrc.Bottom = 180
    brdr.Left = 10
    brdr.Top = 10
    brdr.Right = 290
    brdr.Bottom = 265
    Me.ForeColor = 0
    Me.CurrentX = 70
    Me.CurrentY = 100
    Me.FontSize = 10
    Me.Print "ColorBox ver 2.1"
    Me.Print
    Me.CurrentX = 70
    Me.Print "Copyright © SaifSoft inc."
    Me.CurrentX = 70
    Me.Print "www.saifu.5u.com"
    Me.Print
    Me.CurrentX = 70
    Me.FontSize = 8
    Me.Print "Author: Saifudheen. A. A. "
    Me.CurrentX = 70
    Me.Print "keraleeyan@msn.com"

    DrawEdge hdc, brdr, BDR_SUNKENOUTER Or BDR_RAISEDINNER, BF_RECT
End Sub



Private Sub Text1_Change()
    On Error Resume Next
    Dim cm As CMYK
    If Text1.Text = "" Then Exit Sub
    cm = RGBtoCMYK(Int(Text1.Text), Int(Text2.Text), Int(Text3.Text))
    txtCyan.Text = cm.Cyan
    txtMagenta.Text = cm.Magenta
    txtYellow.Text = cm.Yellow
    txtK.Text = cm.k
    If ChangeByInput = False Then Exit Sub
    If Text1.Text <> Int(Text1.Text) Or Val(Text1.Text) < 0 Or Text1.Text > 255 Then
        Exit Sub
    End If
    SetColors
    ChangeByInput = False
End Sub
Private Sub SetColors()
    'On Error Resume Next

    Dim red As Integer, green As Integer, blue As Integer
    red = Val(Text1.Text)
    green = Val(Text2.Text)
    blue = Val(Text3.Text)
    Dim hs As HSB
    hs = RGBtoHSB(RGB(red, green, blue))
    'Call DrawSlider(SelectedMainPos)
    txtH.Text = Int(hs.Hue)
    txtS.Text = Int(hs.Saturation)
    txtB.Text = Int(hs.Brightness)
    lblSelColor.BackColor = RGB(red, green, blue)
    ReloadColors
    'DrawSlider (SelectedMainPos)
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
ChangeByInput = True

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Text1.Text = "" Or Val(Text1.Text) <> Int(Val(Text1.Text)) Or Val(Text1.Text) < 0 Or Val(Text1.Text) > 255 Then
        MsgBox "Enter an Integer between 0 and 255.", , "ColorBox"
        Cancel = True
        Exit Sub
    End If
    Text1.Text = Val(Text1.Text)

End Sub

Private Sub Text2_Change()
    If Text2.Text = "" Then Exit Sub

    Dim cm As CMYK
    cm = RGBtoCMYK(Int(Text1.Text), Int(Text2.Text), Int(Text3.Text))
    txtCyan.Text = cm.Cyan
    txtMagenta.Text = cm.Magenta
    txtYellow.Text = cm.Yellow
    txtK.Text = cm.k
    If ChangeByInput = False Then Exit Sub
    If Text2.Text <> Int(Text2.Text) Or Val(Text2.Text) < 0 Or Text2.Text > 255 Then
        Exit Sub
    End If
    SetColors
    ChangeByInput = False

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
ChangeByInput = True

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Text2.Text = "" Or Val(Text2.Text) <> Int(Val(Text2.Text)) Or Val(Text2.Text) < 0 Or Val(Text2.Text) > 255 Then
        MsgBox "Enter an Integer between 0 and 255.", , "ColorBox"
        Cancel = True
        Exit Sub
    End If
    Text2.Text = Val(Text2.Text)

End Sub

Private Sub Text3_Change()
    If Text3.Text = "" Then Exit Sub

    Dim cm As CMYK
    If Text3.Text <> Int(Text3.Text) Or Val(Text3.Text) < 0 Or Text3.Text > 255 Then
        Exit Sub
    End If

    cm = RGBtoCMYK(Int(Text1.Text), Int(Text2.Text), Int(Text3.Text))
    txtCyan.Text = cm.Cyan
    txtMagenta.Text = cm.Magenta
    txtYellow.Text = cm.Yellow
    txtK.Text = cm.k
    If ChangeByInput = False Then Exit Sub
    SetColors
    ChangeByInput = False

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
ChangeByInput = True

End Sub

 

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    If Text3.Text = "" Or Val(Text3.Text) <> Int(Val(Text3.Text)) Or Val(Text3.Text) < 0 Or Val(Text3.Text) > 255 Then
        MsgBox "Enter an Integer between 0 and 255.", , "ColorBox"
        Cancel = True
        Exit Sub
    End If
    Text3.Text = Val(Text3.Text)
End Sub




Private Sub txtB_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtB_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrB
    AdjustBrightness txtB.Text
    Exit Sub
ErrB:
    MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
    txtB.SetFocus

End Sub

Sub DrawPicker()
    Me.DrawMode = 6
    Me.FillStyle = 1
    Me.DrawWidth = 1
    Me.DrawStyle = 0
    Me.Circle (SelectedPos.X, SelectedPos.Y), 5
End Sub

Sub DrawSelFrame()
    Dim SelFrame As RECT
    SelFrame.Left = SelectBox.Left - 1
    SelFrame.Top = SelectBox.Top - 1
    SelFrame.Right = SelectBox.Right + 3
    SelFrame.Bottom = SelectBox.Bottom + 3
    DrawEdge Me.hdc, SelFrame, BDR_SUNKENOUTER, BF_RECT
    Me.PSet (-100, -100)
End Sub


Private Sub UpdateColorValues()
        Dim r As Integer, g As Integer, b As Integer
        Dim cl As Long
        cl = Me.Point(SelectedPos.X, SelectedPos.Y)
        lblSelColor.BackColor = cl
        GetRGB cl, r, g, b
        Text1.Text = r
        Text2.Text = g
        Text3.Text = b

End Sub




Private Sub txtB_Validate(Cancel As Boolean)
    If txtB.Text = "" Or Val(txtB.Text) <> Int(Val(txtB.Text)) Or Val(txtB.Text) < 0 Or Val(txtB.Text) > 100 Then
        MsgBox "Enter an Integer between 0 and 100."
        Cancel = True
        Exit Sub
    End If
    txtB.Text = Val(txtB.Text)
    BrightnessEntering = False
End Sub


Private Sub txtH_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtH_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrH
AdjustHue txtH.Text
Exit Sub
ErrH:
    MsgBox "Enter an Integer Value Between 0 and 360", , "ColorBox"
    txtH.SetFocus
End Sub



Private Sub txtH_Validate(Cancel As Boolean)
    If txtH.Text = "" Or Val(txtH.Text) <> Int(Val(txtH.Text)) Or Val(txtH.Text) < 0 Or Val(txtH.Text) > 360 Then
        MsgBox "Enter an Integer between 0 and 360."
        Cancel = True
        Exit Sub
    End If
    txtH.Text = Val(txtH.Text)
  
End Sub

Private Sub txtK_Change()
    If ChangeByInput = False Then Exit Sub
    If txtK.Text <> Int(txtK.Text) Or Val(txtK.Text) < 0 Or txtK.Text > 255 Then
        MsgBox "Enter an Integer between 0 and 255."
        Exit Sub
    End If

End Sub

Private Sub txtK_KeyPress(KeyAscii As Integer)
ChangeByInput = True

End Sub

Private Sub txtK_Validate(Cancel As Boolean)
    If txtK.Text <> Int(txtK.Text) Or Val(txtK.Text) < 0 Or txtK.Text > 255 Then
        MsgBox "Enter an Integer between 0 and 255."
        Cancel = True
        Exit Sub
    End If

End Sub


Private Sub txtS_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtS_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrS
    AdjustSaturation txtS.Text
    Exit Sub
ErrS:
    MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
    txtS.SetFocus

End Sub

Private Sub AdjustHue(ByVal Hue As Single)
        Dim cl As Long
        Dim hs As HSB
        Dim r As Integer, g As Integer, b As Integer
        If Hue > 360 Or Hue < 0 Or Int(Hue) <> Hue Then
            MsgBox "Enter an Integer Value Between 0 and 360", , "ColorBox"
            txtH.SetFocus
            txtH.Text = OldHue
            AdjustHue OldHue
            Exit Sub
        End If
    
        DrawPicker
        DrawSlider SelectedMainPos
        Select Case True
        Case optH.Value
            SelectedMainPos = MainBox.Bottom - Hue * (255 / 360)
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            GetRGB cl, r, g, b
            LoadVariantsHue r, g, b
            UpdateColorValues
        Case optS.Value
            SelectedPos.X = 10 + Hue * (255 / 360)
            hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
            HSBtoRGB hs, r, g, b
            LoadMainSaturation Form1.hdc, r, g, b, Val(txtB.Text) / 100
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        Case optB.Value
            SelectedPos.X = SelectBox.Left + Hue * (255 / 360)
            hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
            cl = Me.Point(SelectedPos.X, SelectedPos.Y)
            HSBtoRGB hs, r, g, b
            LoadMainBrightness Form1.hdc, r, g, b
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        End Select
        DrawPicker
        DrawSlider SelectedMainPos
End Sub

Private Sub AdjustSaturation(ByVal Saturation As Single)
        Dim cl As Long
        Dim r As Integer, g As Integer, b As Integer
        Dim hs As HSB
        If Saturation > 100 Or Saturation < 0 Or Int(Saturation) <> Saturation Then
            MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
            txtS.SetFocus
            txtS.Text = OldSaturation
            AdjustSaturation OldSaturation
            Exit Sub
        End If

        DrawPicker
        DrawSlider SelectedMainPos
        Select Case True
        Case optH.Value
            SelectedPos.X = SelectBox.Left + Saturation * (255 / 100)
            UpdateColorValues
        Case optS.Value
            LoadVariantsSaturation Val(txtS.Text) / 100
            SelectedMainPos = MainBox.Bottom - Saturation * (255 / 100)
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        Case optB.Value
            SelectedPos.Y = SelectBox.Bottom - Saturation * (255 / 100)
            hs.Hue = txtH.Text:  hs.Brightness = 100: hs.Saturation = Val(txtS.Text)
            cl = Me.Point(SelectedPos.X, SelectedPos.Y)
            GetRGB cl, r, g, b
            HSBtoRGB hs, r, g, b
            LoadMainBrightness Form1.hdc, r, g, b
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        End Select
        DrawPicker
        DrawSlider SelectedMainPos
End Sub

Private Sub AdjustBrightness(ByVal Brightness As Single)
        Dim r As Integer, g As Integer, b As Integer
        Dim cl As Long
        Dim hs As HSB
        If Brightness > 100 Or Brightness < 0 Or Int(Brightness) <> Brightness Then
            MsgBox "Enter an Integer Value Between 0 and 100", , "ColorBox"
            txtB.Text = OldBrightness
            txtB.SetFocus
            AdjustBrightness OldBrightness
            Exit Sub
        End If

        DrawPicker
        DrawSlider SelectedMainPos

        Select Case True
        Case optH.Value
            SelectedPos.Y = SelectBox.Bottom - Brightness * (255 / 100)
            UpdateColorValues
        Case optS.Value
            SelectedPos.Y = SelectBox.Bottom - Brightness * (255 / 100)
            hs.Hue = txtH.Text: hs.Saturation = 100: hs.Brightness = 100
            HSBtoRGB hs, r, g, b
            LoadMainSaturation Form1.hdc, r, g, b, Val(txtB.Text) / 100
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
            
        Case optB.Value
            SelectedMainPos = MainBox.Bottom - Brightness * (255 / 100)
            LoadVariantsBrightness Brightness / 100
            cl = Me.Point(SelectedPos.X, SelectedPos.Y)
            GetRGB cl, r, g, b
            cl = Me.Point(MainBox.Left + 5, SelectedMainPos)
            lblSelColor.BackColor = cl
            GetRGB cl, r, g, b
            Text1.Text = r
            Text2.Text = g
            Text3.Text = b
        End Select
        DrawPicker
        DrawSlider SelectedMainPos

End Sub



Private Sub txtS_Validate(Cancel As Boolean)
    If txtS.Text = "" Or Val(txtS.Text) <> Int(Val(txtS.Text)) Or Val(txtS.Text) < 0 Or Val(txtS.Text) > 100 Then
        MsgBox "Enter an Integer between 0 and 100."
        Cancel = True
        Exit Sub
    End If
    txtS.Text = Val(txtS.Text)
    
End Sub




