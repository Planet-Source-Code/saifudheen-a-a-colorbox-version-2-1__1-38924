Attribute VB_Name = "Module1"


'************************ColorBox Version 2.1************************
'Functions module; Color algorithms
'Author : Saifudheen A.A.
'E-mail : keraleeyan@msn.com.
'Suggestions, Votes all are welcome.
'********************************************************************

'Type Declerations
Public Type RGBTRIPLE
    rgbtBlue As Byte
    rgbtGreen As Byte
    rgbtRed As Byte
End Type

Public Type HSB
    Hue As Single
    Saturation As Single
    Brightness As Single
    End Type
    
Public Type CMYK
    Cyan As Integer
    Magenta As Integer
    Yellow As Integer
    k As Integer
    End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAPINFOHEADER       ' 40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmicolors(15) As Long
End Type

Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Public Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry() As PALETTEENTRY
End Type
'API Function Declerations
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetICMMode Lib "gdi32" (ByVal hdc As Long, ByVal n As Long) As Long
Public Declare Function CheckColorsInGamut Lib "gdi32" (ByVal hdc As Long, lpv As Any, lpv2 As Any, ByVal dw As Long) As Long

Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CreatehalfTonePalette Lib "gdi32" Alias "CreateHalftonePalette" (ByVal hdc As Long) As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

' Constants
Public Const BitsPixel = 12
Public Const Planes = 14

Public Const BDR_RAISEDINNER = &H4

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2

Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_LEFT = &H1
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Public Const ICM_ON = 2
Public Const ICM_OFF = 1
Public Const ICM_QUERY = 3

Public Const DIB_RGB_COLORS = 0 '  color table in RGBs
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices

'Variables
Public pMode As Integer
Dim lpbmINFO As BITMAPINFO
Public lpBI As BITMAPINFO
Public m_Color As Long
Public mOldColor As Long
Public SelectBox As RECT
Public MainBox As RECT
Public Preset() As RECT
Public SelectedPos As POINTAPI
Public SelectedMainPos As Single
Public cPaletteIndex As Integer
Public optNotClicked As Boolean
Public svdColor() As Long
Public m_WebColors As Boolean
Public LastColor As Long




Sub LoadVariantsHue(red As Integer, green As Integer, blue As Integer)
  
    'On Error Resume Next
    Dim St As Long
    
    Dim X As Integer, Y As Integer
    Dim sDc As Long
    Dim K1 As Double, K2 As Double, K3 As Double
    K1 = red / 255
    K2 = green / 255
    K3 = blue / 255
    With Form1
        .DrawWidth = 1
        .DrawMode = 13
        
        Dim M1    As Double, M2     As Double, M3     As Double
        Dim J1    As Double, J2     As Double, J3     As Double
        Dim YMax As Byte
        Dim shdBitmap(0 To 196608) As Byte  '256 ^ 2 * 3
        Dim L As Long
        Dim bpos As Long
        Dim count As Long
        bpos = 0
        count = 0
        
        With lpBI.bmiHeader
            .biHeight = 256
            .biWidth = 256
        End With
        
        On Error Resume Next
        For Y = 255 To 0 Step -1
                 M1 = red - Y * K1
                 M2 = green - Y * K2
                 M3 = blue - Y * K3
                 YMax = 255 - Y
                 J1 = (YMax - M1) / 255
                 J2 = (YMax - M2) / 255
                 J3 = (YMax - M3) / 255
            For X = 255 To 0 Step -1
                If m_WebColors Then
                    shdBitmap(bpos) = CInt((M3 + X * J3) / &H33) * &H33
                    shdBitmap(bpos + 1) = CInt((M2 + X * J2) / &H33) * &H33
                    shdBitmap(bpos + 2) = CInt((M1 + X * J1) / &H33) * &H33
                Else
                    shdBitmap(bpos) = M3 + X * J3   'Blue
                    shdBitmap(bpos + 1) = M2 + X * J2 'Green
                    shdBitmap(bpos + 2) = M1 + X * J1 'Red
                End If
                bpos = bpos + 3
            Next X
        Next Y
        
    BltBitmap Form1.hdc, shdBitmap, 10, 10, 256, 256, True
       
        
    End With
    Form1.DrawSelFrame
End Sub




Sub LoadVariantsBrightness(ByVal Brightness As Single)
Dim OldP As POINTAPI
Dim v As Integer
On Error Resume Next
Dim H, M As Single
Dim a As Integer, b As Integer, c As Integer, D As Integer, E As Integer, f As Integer
Dim sDc As Long
Dim Color As Long
Dim red As Integer, green As Integer, blue As Integer
H = SelectBox.Bottom - SelectBox.Top
M = H / 6
a = M
b = 2 * M
c = 3 * M
D = 4 * M
E = 5 * M
f = 6 * M
Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3

With Form1
    .DrawMode = 13
    sDc = .hdc
End With

With lpBI.bmiHeader
    .biHeight = 256
    .biWidth = 256
End With

    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * a  'These are the common Terms taken out from the For Loop for efficciency
    Mcc = 255 + 6 * c  ' ""
    Mee = 255 + 6 * E  ' ""
    Dim pos As Long
    pos = 0
    
Dim X  As Integer, Y As Integer
For Y = 255 To 0 Step -1
        MV = 1 - Y / 255 ' ""
    '1
        For X = 0 To a
            v = X * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 0) = CInt(Kc * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = Y * Brightness
                bBitmap(pos + 0) = Kc * Brightness
            End If
            pos = pos + 3
        Next X
    '2
        For X = a + 1 To b
            v = Maa - 6 * X ' 255 - (X - A) * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 1) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 0) = 255 * Brightness
            Else
                bBitmap(pos + 2) = Kc * Brightness
                bBitmap(pos + 1) = Y * Brightness
                bBitmap(pos + 0) = 255 * Brightness
            End If
            pos = pos + 3
        Next X
     '3
        For X = b + 1 To c
            v = (X - b - 1) * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 1) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 0) = 255 * Brightness
            Else
                bBitmap(pos + 2) = Y * Brightness
                bBitmap(pos + 1) = Kc * Brightness
                bBitmap(pos + 0) = 255 * Brightness
            End If
            pos = pos + 3
        Next X
     '4
        For X = c + 1 To D
            v = Mcc - 6 * X
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Y * Brightness / &H33) * &H33
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = CInt(Kc * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = Y * Brightness
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = Kc * Brightness
            End If
            pos = pos + 3
        Next X
    '5
        For X = D + 1 To E
            v = (X - D - 1) * 6
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = CInt(Y * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = Kc * Brightness
                bBitmap(pos + 1) = 255 * Brightness
                bBitmap(pos + 0) = Y * Brightness
            End If
            pos = pos + 3
        Next X
    '6
        For X = E + 1 To f
            v = Mee - 6 * X
            Kc = v * MV + Y
            If m_WebColors Then
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = CInt(Kc * Brightness / &H33) * &H33
                bBitmap(pos + 0) = CInt(Y * Brightness / &H33) * &H33
            Else
                bBitmap(pos + 2) = 255 * Brightness
                bBitmap(pos + 1) = Kc * Brightness
                bBitmap(pos + 0) = Y * Brightness
            End If
            pos = pos + 3
        Next X
       
Next Y

    BltBitmap Form1.hdc, bBitmap, 265, 10, -256, 256, True
    Form1.DrawSelFrame
    
End Sub


Sub LoadVariantsSaturation(ByVal Saturation As Single)
Dim OldP As POINTAPI
Dim v As Integer
On Error Resume Next
Dim H, M As Single
Dim X As Integer, Y As Integer
Dim a As Integer, b As Integer, c As Integer, D As Integer, E As Integer, f As Integer
Dim sDc As Long
Dim Color As Long
Dim red As Integer, green As Integer, blue As Integer
Dim bBitmap(0 To 256 ^ 2 * 3) As Byte '256 ^ 2 * 3
Dim cpos As Long
cpos = 0
H = SelectBox.Bottom - SelectBox.Top
M = H / 6
a = M
b = 2 * M
c = 3 * M
D = 4 * M
E = 5 * M
f = 6 * M

'Form1.DrawMode = 6
'Form1.Circle (SelectedPos.x, SelectedPos.y), 5  'Erases Previous Circle

With Form1
    .DrawWidth = 1
    .DrawMode = 13
    sDc = .hdc
End With
    Dim Maa As Double, Mcc As Double, Mee As Double
    Dim MV As Double
    Dim Kc As Integer
    Dim YPos As Long
    Maa = 255 + 6 * a  'These are the common Terms taken out from the For Loop for efficiency
    Mcc = 255 + 6 * c  ' ""
    Mee = 255 + 6 * E  ' ""
    
For Y = 255 To 0 Step -1
        MV = 1 - Y / 255  ' ""
        YPos = SelectBox.Top + Y
    '1
        For X = 0 To a
            v = X * 6
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 0) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
            Else
                bBitmap(pos + 2) = (255 - Y)
                bBitmap(pos + 1) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 0) = Kc + (255 - Y - Kc) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
    '2
        For X = a + 1 To b
            v = Maa - 6 * X
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) / &H33) * &H33
            Else
                bBitmap(pos + 2) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 1) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 0) = (255 - Y)
            End If
            pos = pos + 3
        Next X
     '3
        For X = b + 1 To c
            v = (X - b - 1) * 6
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 1) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) / &H33) * &H33
            Else
                bBitmap(pos + 2) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 1) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 0) = (255 - Y)
            End If
            pos = pos + 3
        Next X
     '4
        For X = c + 1 To D
            v = Mcc - 6 * X
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 0) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
            Else
                bBitmap(pos + 2) = (255 - Y) * (1 - Saturation)
                bBitmap(pos + 1) = 255 - Y
                bBitmap(pos + 0) = Kc + (255 - Y - Kc) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
    '5
        For X = D + 1 To E
            v = (X - D - 1) * 6
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 1) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
            Else
                bBitmap(pos + 2) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 1) = 255 - Y
                bBitmap(pos + 0) = (255 - Y) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
    '6
        For X = E + 1 To f
            v = Mee - 6 * X
            Kc = v * MV
            If m_WebColors Then
                bBitmap(pos + 2) = CInt((255 - Y) / &H33) * &H33
                bBitmap(pos + 1) = CInt((Kc + (255 - Y - Kc) * (1 - Saturation)) / &H33) * &H33
                bBitmap(pos + 0) = CInt((255 - Y) * (1 - Saturation) / &H33) * &H33
            Else
                bBitmap(pos + 2) = 255 - Y
                bBitmap(pos + 1) = Kc + (255 - Y - Kc) * (1 - Saturation)
                bBitmap(pos + 0) = (255 - Y) * (1 - Saturation)
            End If
            pos = pos + 3
        Next X
       
Next Y
    BltBitmap Form1.hdc, bBitmap, 265, 10, -256, 256, True
    Form1.DrawSelFrame
    'Form1.DrawMode = 6
    'Form1.Circle (SelectedPos.x, SelectedPos.y), 5 'Refresh Circle

End Sub

Sub GetRGB(ByRef cl As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
    Dim c As Long
    c = cl
    red = c Mod &H100
    c = c \ &H100
    green = c Mod &H100
    c = c \ &H100
    blue = c Mod &H100
End Sub

Sub DrawSlider(ByVal Position As Integer)
    Form1.DrawMode = 6
    Form1.DrawWidth = 2
    Form1.Line (MainBox.Right + 2, Position)-(MainBox.Right + 5, Position)
    Form1.Line (MainBox.Left - 2, Position)-(MainBox.Left - 5, Position)
    Form1.DrawWidth = 1
End Sub

Sub LoadSafePalette()
Form1.FillStyle = 0
Form1.DrawMode = 13
Form1.DrawWidth = 1
On Error Resume Next
Dim i, j, k As Integer
Dim L As Long
Dim count As Integer
Dim Plt As Long
Dim ret As Long
Dim br As Long
Dim pal As Long, oldpal As Long
pal = CreatehalfTonePalette(Form1.hdc)
oldpal = SelectPalette(Form1.hdc, pal, 0)
RealizePalette (Form1.hdc)

For i = 0 To &HFF Step &H33
    For j = 0 To &HFF Step &H33
        For k = 0 To &HFF Step &H33
            count = count + 1
            DrawSafeColor Preset(count), i, j, k
        Next k
    Next j
Next i


For i = 217 To 224
    Form1.FillColor = 0
    Rectangle Form1.hdc, Preset(i).Left, Preset(i).Top, Preset(i).Right, Preset(i).Bottom
Next i

SelectPalette Form1.hdc, oldpal, 0
DeleteObject pal

Form1.DrawSafePicker cPaletteIndex, False
Dim r As Integer, g As Integer, b As Integer
Form1.GetSafeColor cPaletteIndex, r, g, b
'Form1.lblSelColor.BackColor = RGB(r, g, b)
End Sub

Sub LoadCustomColors()
    Dim FileHandle As Integer
    Dim i As Integer
    Dim strColor As String
    On Error Resume Next
    FileHandle = FreeFile()
    ReDim svdColor(0 To 224)
    Open App.Path & "/usercolors.cps" For Input As #FileHandle
    i = 0
    Form1.Cls
    Form1.PrintLastColor
    Form1.FillStyle = 0
    Form1.DrawMode = 13
    Form1.DrawWidth = 1
    For i = 0 To 224
        Line Input #FileHandle, strColor
        svdColor(i) = Val(strColor)
        Form1.ForeColor = vbBlack 'svdColor(i)
        Form1.FillColor = svdColor(i)
        Rectangle Form1.hdc, Preset(i).Left, Preset(i).Top, Preset(i).Right, Preset(i).Bottom
    Next i
    Close #FileHandle
    Form1.DrawSafePicker cPaletteIndex, False
    Form1.PSet (-100, -100)
End Sub

Sub SaveCustomColors()
    Dim FileHandle As Integer
    Dim i As Integer
    On Error Resume Next
    FileHandle = FreeFile()
    Open App.Path & "/usercolors.cps" For Output As #FileHandle
    For i = 0 To 224
        Print #FileHandle, svdColor(i)
    Next i
    Close #FileHandle
  
End Sub



Public Sub LoadMainSaturation(ByVal hdc As Long, ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer, ByVal Brightness As Single)

    Dim r As Integer, g As Integer, b As Integer
    Dim bBitmap(17 * 256 * 3) As Byte
    Dim pos As Integer
    Dim X As Long, Y As Long
    Dim f As Single
    For Y = 0 To 255
        f = 1 - (Y / 255)
        r = red * f + Y
        g = green * f + Y
        b = blue * f + Y

        For X = 0 To 15
            If m_WebColors Then
                bBitmap(pos) = CInt(b * Brightness / &H33) * &H33
                bBitmap(pos + 1) = CInt(g * Brightness / &H33) * &H33
                bBitmap(pos + 2) = CInt(r * Brightness / &H33) * &H33
            Else
                bBitmap(pos) = b * Brightness
                bBitmap(pos + 1) = g * Brightness
                bBitmap(pos + 2) = r * Brightness
            End If
            pos = pos + 3
        Next X
    Next Y

    BltBitmap hdc, bBitmap, 277, 265, 15, -256, True
    DrawMainFrame
End Sub

Public Sub LoadMainHue()
    Dim OldP As POINTAPI
    Dim v As Integer
    On Error Resume Next
    Dim H As Single, M As Single
    Dim a As Single, b As Single, c As Single, D As Single, E As Single, f As Single
    Dim Ratio As Single
    
    H = SelectBox.Bottom - SelectBox.Top
    M = H / 6
    a = M
    b = 2 * M
    c = 3 * M
    D = 4 * M
    E = 5 * M
    f = 6 * M
    Dim sBitmap(0 To 16 * 256 * 3) As Byte            '256 ^ 2 * 3
    Dim cpos  As Long
    With lpBI.bmiHeader
        .biHeight = 256
        .biWidth = 15
    End With

    cpos = 0
        For Y = 0 To Int(a)
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = CInt(Y * 6 / &H33) * &H33
                Else
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = Y * 6
                End If
                cpos = cpos + 3
            Next j
        Next Y
    '2
                
        For Y = Int(a) + 1 To Int(b)
            v = 255 - (Y - a) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = 255
                Else
                    sBitmap(cpos + 2) = v
                    sBitmap(cpos + 1) = 0
                    sBitmap(cpos + 0) = 255
                End If
                cpos = cpos + 3
            Next j
            
        Next Y
     '3
         
        For Y = Int(b) + 1 To Int(c)
            v = (Y - b) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 0) = 255
                Else
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = v
                    sBitmap(cpos + 0) = 255
                End If
                cpos = cpos + 3
            Next j
            
        Next Y
     '4
        For Y = Int(c) + 1 To Int(D)
            v = 255 - (Y - c) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = CInt(v / &H33) * &H33
                Else
                    sBitmap(cpos + 2) = 0
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = v
                End If
                cpos = cpos + 3
            Next j
        Next Y
    '5
        For Y = Int(D) + 1 To Int(E)
            v = (Y - D) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = 0
                Else
                    sBitmap(cpos + 2) = v
                    sBitmap(cpos + 1) = 255
                    sBitmap(cpos + 0) = 0
                End If
                cpos = cpos + 3
            Next j
            
        Next Y
    '6
        For Y = Int(E) + 1 To Int(f)
            v = 255 - (Y - E) * 6
            For j = 1 To 16
                If m_WebColors Then
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = CInt(v / &H33) * &H33
                    sBitmap(cpos + 0) = 0
                Else
                    sBitmap(cpos + 2) = 255
                    sBitmap(cpos + 1) = v
                    sBitmap(cpos + 0) = 0
                End If
                cpos = cpos + 3
            Next j
        Next Y
        BltBitmap Form1.hdc, sBitmap, MainBox.Left, MainBox.Bottom, 15, -256, True
        DrawMainFrame
        'Form1.DrawPicker
End Sub


Public Sub LoadMainBrightness(ByVal hdc As Long, ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer)
    Dim r As Integer, g As Integer, b As Integer
    Dim bBitmap(17 * 256 * 3) As Byte
    Dim pos As Integer
    Dim X As Long, Y As Long
    
    For Y = 0 To 255
        r = red - red * Y / 255
        g = green - green * Y / 255
        b = blue - blue * Y / 255
        For X = 0 To 15
            If m_WebColors Then
                bBitmap(pos) = CInt(b / &H33) * &H33
                bBitmap(pos + 1) = CInt(g / &H33) * &H33
                bBitmap(pos + 2) = CInt(r / &H33) * &H33
            Else
                bBitmap(pos) = b
                bBitmap(pos + 1) = g
                bBitmap(pos + 2) = r
            End If
            pos = pos + 3
        Next X
    Next Y
    
    BltBitmap hdc, bBitmap, 277, 265, 15, -256, True
    DrawMainFrame
End Sub
   
Public Function HSBtoRGB(hs As HSB, ByRef r As Integer, ByRef g As Integer, ByRef b As Integer)
    Dim i As Integer
    Dim f As Single, p As Single, q As Single, t As Single
    hs.Saturation = hs.Saturation * 255 / 100
    hs.Brightness = hs.Brightness * 255 / 100
    If (hs.Saturation = 0) Then
        ' achromatic (grey)
        r = hs.Brightness
        g = r
        b = r
        Exit Function
    End If
    
    hs.Hue = hs.Hue / 60         ' sector 0 to 5
    i = Int(hs.Hue)
    f = hs.Hue - i         ' factorial part of hs.Hue
    p = hs.Brightness * (1 - hs.Saturation / 255)
    q = hs.Brightness * (1 - (hs.Saturation / 255) * f)
    t = hs.Brightness * (1 - (hs.Saturation / 255) * (1 - f))
    Select Case i
        Case 0
            r = hs.Brightness
            g = t
            b = p
        Case 1
            r = q
            g = hs.Brightness
            b = p
        Case 2
            r = p
            g = hs.Brightness
            b = t
        Case 3
            r = p
            g = q
            b = hs.Brightness
        Case 4
            r = t
            g = p
            b = hs.Brightness
        Case Else        ' case 5:
            r = hs.Brightness
            g = p
            b = q
        End Select
    
End Function

Public Function RGBtoHSB(ByVal Color As Long) As HSB
    Dim LargestValue As Integer
    Dim SmallestValue As Integer
    Dim red  As Integer, green As Integer, blue As Integer
    Dim RedRatio As Single, GreenRatio As Single, BlueRatio As Single
    GetRGB Color, red, green, blue
    LargestValue = IIf(red >= green, red, green)
    LargestValue = IIf(LargestValue >= blue, LargestValue, blue)
    SmallestValue = IIf(red <= green, red, green)
    SmallestValue = IIf(SmallestValue <= blue, SmallestValue, blue)
    RGBtoHSB.Brightness = LargestValue * 100 / 255
    If LargestValue <> 0 Then
        RGBtoHSB.Saturation = 100 - (SmallestValue * 100 / LargestValue)
    Else
        RGBtoHSB.Saturation = 0
    End If
    If RGBtoHSB.Saturation = 0 Then
        RGBtoHSB.Hue = 0
    Else
        RedRatio = (LargestValue - red) / (LargestValue - SmallestValue)
        GreenRatio = (LargestValue - green) / (LargestValue - SmallestValue)
        BlueRatio = (LargestValue - blue) / (LargestValue - SmallestValue)
        Select Case LargestValue
        Case red
            RGBtoHSB.Hue = BlueRatio - GreenRatio
        Case green
            RGBtoHSB.Hue = (2 + RedRatio) - BlueRatio
        Case blue
            RGBtoHSB.Hue = (4 + GreenRatio) - RedRatio
        End Select
        RGBtoHSB.Hue = RGBtoHSB.Hue * 60
        If RGBtoHSB.Hue < 0 Then
            RGBtoHSB.Hue = RGBtoHSB.Hue + 360
        End If
    End If
    

End Function

Sub DrawMainFrame()
    Dim MainFrame As RECT
     
    MainFrame.Left = MainBox.Left - 1
    MainFrame.Top = MainBox.Top - 1
    MainFrame.Right = MainBox.Right + 1
    MainFrame.Bottom = MainBox.Bottom + 3
    DrawEdge Form1.hdc, MainFrame, BDR_SUNKENOUTER, BF_RECT
    Form1.PSet (-100, -100)
End Sub

Public Function RGBtoCMYK(ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer) As CMYK
    With RGBtoCMYK
        .Cyan = 255 - red
        .Magenta = 255 - green
        .Yellow = 255 - blue
        .k = IIf(.Cyan < .Magenta, .Cyan, .Magenta)
        If .Yellow < .k Then
            .k = .Yellow
        End If
        If .k > 0 Then
            .Cyan = .Cyan - .k
            .Magenta = .Magenta - .k
            .Yellow = .Yellow - .k
        End If
        Dim MinColor   As Integer
        MinColor = IIf(.Cyan < .Magenta, .Cyan, .Magenta)
        MinColor = IIf(.Yellow < MinColor, .Yellow, MinColor)
        MinColor = IIf((MinColor + .k) > 255, 255 - .k, MinColor)
        .Cyan = (.Cyan - MinColor)
        .Magenta = (.Magenta - MinColor)
        .Yellow = (.Yellow - MinColor)
        .k = (.k + MinColor)
    End With
End Function



Public Sub DrawSafeColor(sPos As RECT, ByVal red As Integer, ByVal green As Integer, ByVal blue As Integer)

    Dim bBitmap(3 * 16 * 16) As Byte
    Dim pos As Integer
    Dim X As Long, Y As Long
    Dim Width As Integer
    Dim Height As Integer
    Width = sPos.Bottom - sPos.Top
    Height = sPos.Bottom - sPos.Top
    For Y = 0 To Height
        For X = 0 To Width
                bBitmap(pos) = blue
                bBitmap(pos + 1) = green
                bBitmap(pos + 2) = red
                pos = pos + 3
        Next X
    Next Y
    BltBitmap Form1.hdc, bBitmap, sPos.Left, sPos.Top, Width, Height, False


End Sub

Public Sub BltBitmap(ByVal hdc As Long, bmptr() As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal CreatehfTPalette As Boolean)
        Dim lpBI As BITMAPINFO
        lpBI.bmiHeader.biBitCount = 24
        lpBI.bmiHeader.biCompression = BI_RGB
        lpBI.bmiHeader.biWidth = Abs(Width)
        lpBI.bmiHeader.biHeight = Abs(Height)
        lpBI.bmiHeader.biPlanes = 1
        lpBI.bmiHeader.biSize = 40
        If CreatehfTPalette Then
            Dim pal As Long, oldpal As Long
            pal = CreatehalfTonePalette(hdc)
            oldpal = SelectPalette(hdc, pal, 0)
            RealizePalette (hdc)
        End If
        StretchDIBits hdc, X, Y, Width, Height, 0, 0, Abs(Width), Abs(Height), bmptr(0), lpBI, DIB_RGB_COLORS, vbSrcCopy
        If CreatehfTPalette Then
            SelectPalette hdc, oldpal, 0
            DeleteObject pal
        End If
End Sub
