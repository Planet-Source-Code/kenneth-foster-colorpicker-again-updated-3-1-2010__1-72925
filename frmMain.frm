VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "cp2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   5145
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGraybar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   990
      Picture         =   "frmMain.frx":28D0
      ScaleHeight     =   270
      ScaleWidth      =   4035
      TabIndex        =   21
      Top             =   5715
      Width           =   4065
   End
   Begin VB.PictureBox picColorbar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   990
      Picture         =   "frmMain.frx":36AB
      ScaleHeight     =   270
      ScaleWidth      =   4035
      TabIndex        =   20
      Top             =   5370
      Width           =   4065
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3435
      Width           =   1545
   End
   Begin VB.PictureBox picPickBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   750
      Picture         =   "frmMain.frx":4135
      ScaleHeight     =   270
      ScaleWidth      =   4035
      TabIndex        =   4
      Top             =   2670
      Width           =   4065
      Begin VB.Shape Shape1 
         Height          =   300
         Left            =   3990
         Top             =   -15
         Width           =   15
      End
   End
   Begin VB.CheckBox chkGrayscale 
      BackColor       =   &H0062ADE8&
      Caption         =   "Gray Scale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1725
      TabIndex        =   3
      Top             =   3090
      Width           =   1260
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   4680
   End
   Begin VB.PictureBox PicHiddenData 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      Picture         =   "frmMain.frx":4BBF
      ScaleHeight     =   15
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   -2000
      Visible         =   0   'False
      Width           =   9611
   End
   Begin Project1.ctlPickColor cp1 
      Height          =   1395
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   2461
      BkgdColor       =   12640511
   End
   Begin Project1.ucKF_Slider Hs1 
      Height          =   150
      Left            =   750
      TabIndex        =   2
      Top             =   1890
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   265
      Min             =   -255
      Max             =   255
      Value           =   0
      SliderColor     =   16777215
      CenterTick      =   -1  'True
      SliderShape     =   2
      SliderBackColor =   6467048
   End
   Begin Project1.CandyButton cmdGetPixel 
      Height          =   390
      Left            =   4740
      TabIndex        =   17
      Top             =   3975
      Width           =   1035
      _ExtentX        =   1879
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Get Pixel"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      ColorScheme     =   0
   End
   Begin Project1.CandyButton cmdExit 
      Height          =   375
      Left            =   4740
      TabIndex        =   18
      Top             =   4425
      Width           =   1020
      _ExtentX        =   1879
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "EXIT"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   8421631
      ColorButtonUp   =   255
      ColorButtonDown =   255
      BorderBrightness=   0
      ColorBright     =   12632319
      ColorScheme     =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Click to transfer >> values"
      Height          =   720
      Left            =   3570
      TabIndex        =   22
      Top             =   2070
      Width           =   750
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Pick Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2100
      TabIndex        =   19
      Top             =   2475
      Width           =   1320
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4035
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4035
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      X1              =   2850
      X2              =   2850
      Y1              =   3420
      Y2              =   4395
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "VB Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   300
      TabIndex        =   16
      Top             =   3480
      Width           =   870
   End
   Begin VB.Label lblInvLongValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2910
      TabIndex        =   15
      Top             =   3780
      Width           =   1035
   End
   Begin VB.Label lblInvRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RGB(0,0,0)"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2910
      TabIndex        =   14
      Top             =   4095
      Width           =   1620
   End
   Begin VB.Label lblInvColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inverse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2910
      TabIndex        =   13
      Top             =   3435
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on a box to send to Clipboard"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   1140
      TabIndex        =   12
      Top             =   4335
      Width           =   2520
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RGB Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   4110
      Width           =   1095
   End
   Begin VB.Label lblRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RGB(0,0,0)"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1155
      TabIndex        =   10
      Top             =   4095
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Long Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   3795
      Width           =   1095
   End
   Begin VB.Label lblLongValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1155
      TabIndex        =   8
      Top             =   3780
      Width           =   1035
   End
   Begin VB.Label lblBrightColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   4350
      TabIndex        =   6
      Top             =   2190
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2325
      TabIndex        =   5
      Top             =   2040
      Width           =   870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'by Ken Foster
'February 2010

Dim rgbvalue As Long
Dim pt As POINTAPI
Dim activewindow As String
Dim R As Long
Dim G As Long
Dim B As Long
Dim activetxt As String
Dim PosX
Dim PosY

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim tt As Boolean

'Stay on Top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const conHwndTopmost = -1
Private Const conSwpNoActivate = &H10
Private Const conSwpShowWindow = &H40

'Get pixel declares
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

'Fast binary Data
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Dim PicInfo As BITMAP

Private Type BITMAP
 bmType As Long
 bmWidth As Long
 bmHeight As Long
 bmWidthBytes As Long
 bmPlanes As Integer
 bmBitsPixel As Integer
 bmBits As Long
End Type

Private Sub Form_Load()
Dim Region As Long
Dim ByteCtr As Long
Dim ByteData(2559) As Byte

frmMain.BackColor = RGB(243, 181, 54)
cp1_ColorPicked
'stay on top
SetWindowPos hWnd, conHwndTopmost, 0, 0, 408, 340, conSwpNoActivate Or conSwpShowWindow
frmMain.SetFocus
tt = False
ByteCtr = 2560

'Get the Data
GetObject PicHiddenData.Image, Len(PicInfo), PicInfo
GetBitmapBits PicHiddenData.Image, ByteCtr, ByteData(0)

'Shape The Form
Region = ExtCreateRegion(ByVal 0&, ByteCtr, ByteData(0))
SetWindowRgn Me.hWnd, Region, True
Hs1.SliderBackColor = &H62ADE8    'set backcolor
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PosX = X  'move form
    PosY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then   'move form
        Me.Left = Me.Left + (X - PosX)
        Me.Top = Me.Top + (Y - PosY)
    End If
End Sub

Private Sub Form_Resize()
   lblLongValue.Caption = rgbvalue
   lblBrightColor.BackColor = cp1.Color
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub cmdGetPixel_Click()
    activetxt = GetActiveWindow
    If tt = False Then
        tt = True
        Timer1.Interval = 50
        Timer1.Enabled = True
    Else
        tt = False
        Timer1.Interval = 0
        Timer1.Enabled = False
    End If
End Sub

Private Sub chkGrayscale_Click()
   If chkGrayscale.Value = Checked Then
      picPickBar.Picture = picGraybar.Picture
   Else
      picPickBar.Picture = picColorbar.Picture
   End If
   picPickBar_MouseDown 0, 0, -80, 0   'hide shape1
   If chkGrayscale.Value = Unchecked Then
       If Shape1.Bordercolor > 7500402 Then Shape1.Bordercolor = vbWhite
    Else
       Shape1.Bordercolor = lblInvColor.BackColor
    End If
    tt = False
End Sub

Private Sub Hs1_Changed()
    lblBrightColor.BackColor = AdjustBrightness(cp1.Color, Hs1.Value)
End Sub

Private Sub lblBrightColor_Click()
Dim bR As Long, bG As Long, bB As Long

   GetRGB lblBrightColor.BackColor, bR, bG, bB
   cp1.setRedValue bR
   cp1.setGreenValue bG
   cp1.setBlueValue bB
End Sub

Private Sub lblInvLongValue_Click()
   Clipboard.Clear
   Clipboard.SetText lblInvLongValue.Caption
End Sub

Private Sub lblInvRGB_Click()
   Clipboard.Clear
   Clipboard.SetText lblInvRGB.Caption
End Sub

Private Sub lblRGB_Click()
   Clipboard.Clear
   Clipboard.SetText lblRGB.Caption
End Sub

Private Sub lblLongValue_Click()
   Clipboard.Clear
   Clipboard.SetText lblLongValue.Caption
End Sub

Private Sub cp1_ColorPicked()
    If tt = False Then
       rgbvalue = cp1.Color
       lblLongValue.Caption = rgbvalue
       GetRGB rgbvalue, R, G, B
       lblRGB.Caption = "RGB(" & R & "," & G & "," & B & ")"
    'inverse color
       lblInvColor.ForeColor = cp1.Color
       lblInvColor.BackColor = RGB(cp1.getInvRedValue, cp1.getInvGreenValue, cp1.getInvBlueValue)
       lblInvRGB.Caption = "RGB(" & cp1.getInvRedValue & "," & cp1.getInvGreenValue & "," & cp1.getInvBlueValue & ")"
       lblInvLongValue.Caption = CLng(lblInvColor.BackColor)
       Text1.Text = "&H00" & GetHex(B) & GetHex(G) & GetHex(R) & "&"
    End If
    Hs1.Value = 0
    
    lblBrightColor.BackColor = cp1.Color
End Sub

Private Sub picPickBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    Dim stdout As Long
    GetCursorPos pt
    tt = True
    rgbvalue = GetPixel(GetDC(stdout), pt.X, pt.Y)
    lblLongValue.Caption = rgbvalue
    GetRGB rgbvalue, R, G, B
    lblRGB.Caption = "RGB(" & R & "," & G & "," & B & ")"
    cp1.setRedValue (R)
    cp1.setGreenValue (G)
    cp1.setBlueValue (B)
     'inverse color
    lblInvColor.ForeColor = cp1.Color
    lblInvColor.BackColor = RGB(cp1.getInvRedValue, cp1.getInvGreenValue, cp1.getInvBlueValue)
    lblInvRGB.Caption = "RGB(" & cp1.getInvRedValue & "," & cp1.getInvGreenValue & "," & cp1.getInvBlueValue & ")"
    lblInvLongValue.Caption = CLng(lblInvColor.BackColor)
    Text1.Text = "&H00" & GetHex(B) & GetHex(G) & GetHex(R) & "&"
    If chkGrayscale.Value = Unchecked Then
       If Shape1.Bordercolor > 7500402 Then Shape1.Bordercolor = vbWhite
    Else
       Shape1.Bordercolor = lblInvColor.BackColor
    End If
    Shape1.Left = X
End Sub

Private Sub picPickBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If chkGrayscale.Value = Checked Then
       If Shape1.Bordercolor > 7500402 Then Shape1.Bordercolor = vbWhite
   Else
       Shape1.Bordercolor = lblInvColor.BackColor
   End If
   tt = False
End Sub

Private Sub Text1_Click()
   Clipboard.Clear
   Clipboard.SetText Text1.Text
End Sub

Private Sub GetRGB(cl As Long, Red As Long, Green As Long, Blue As Long)
    Dim C As Long
    C = cl
    Red = C Mod &H100
    C = C \ &H100
    Green = C Mod &H100
    C = C \ &H100
    Blue = C Mod &H100
End Sub

Private Function GetHex(intVal As Long) As String
    Dim strHex As String
    strHex = Hex(intVal)
    If Len(strHex) = 1 Then strHex = "0" & strHex
    GetHex = strHex
End Function

Private Sub Timer1_Timer()
Dim stdout As Long
Dim sColor As Double

On Error Resume Next
activewindow = GetActiveWindow
If activewindow <> activetxt Then
    Timer1.Interval = 0
    Timer1.Enabled = False
    tt = False
Else
   On Error GoTo leave
    GetCursorPos pt
    rgbvalue = GetPixel(GetDC(stdout), pt.X, pt.Y)
    lblLongValue.Caption = rgbvalue
    GetRGB rgbvalue, R, G, B
    lblRGB.Caption = "RGB(" & R & "," & G & "," & B & ")"
    cp1.setRedValue (R)
    cp1.setGreenValue (G)
    cp1.setBlueValue (B)
     'inverse color
    lblInvColor.ForeColor = cp1.Color
    lblInvColor.BackColor = RGB(cp1.getInvRedValue, cp1.getInvGreenValue, cp1.getInvBlueValue)
    lblInvRGB.Caption = "RGB(" & cp1.getInvRedValue & "," & cp1.getInvGreenValue & "," & cp1.getInvBlueValue & ")"
    lblInvLongValue.Caption = CLng(lblInvColor.BackColor)
    Text1.Text = "&H00" & GetHex(B) & GetHex(G) & GetHex(R) & "&"
End If
    frmMain.SetFocus
leave:
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Function AdjustBrightness(ByVal Color As Long, ByVal Amount As Single) As Long
    On Error Resume Next
    Dim R0 As Long, G0 As Long, B0 As Long
    Dim R1 As Long, G1 As Long, B1 As Long

    'get red, green, and blue values
    GetRGB2 R0, G0, B0, Color

    'add/subtract the amount to/from the ori
    '     ginal RGB values
    R1 = SetBound(R0 + Amount, 0, 255)
    G1 = SetBound(G0 + Amount, 0, 255)
    B1 = SetBound(B0 + Amount, 0, 255)

    'convert RGB back to Long value
    AdjustBrightness = RGB(R1, G1, B1)
End Function

Private Function SetBound(ByVal Num As Single, ByVal MinNum As Single, ByVal MaxNum As Single) As Single

    If Num < MinNum Then
        'if less that min value make it the min
        '     value
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        'if more than max value make it the max
        '     value
        SetBound = MaxNum
    Else
        'if between min and max then leave it al
        '     one
        SetBound = Num
    End If
End Function

Private Sub GetRGB2(R As Long, G As Long, B As Long, ByVal Color As Long)
    Dim TempValue As Long

    'First translate the color from a long v
    '     alue to a short value
    TranslateColor Color, 0, TempValue

    'Calculate the red, green, and blue valu
    '     es from the short value
    R = TempValue And &HFF&
    G = (TempValue And &HFF00&) / 2 ^ 8
    B = (TempValue And &HFF0000) / 2 ^ 16
End Sub

