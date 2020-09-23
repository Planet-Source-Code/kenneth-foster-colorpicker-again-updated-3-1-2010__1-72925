VERSION 5.00
Begin VB.UserControl ctlPickColor 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   ToolboxBitmap   =   "PickColor.ctx":0000
End
Attribute VB_Name = "ctlPickColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Note: this control requires cMemoryBmp Class to work
'original code by Jim Benvenuti
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
Const PS_SOLID = 0
Const SYSTEM_FONT = 13
Const DT_TOP = &H0
Const DT_LEFT = &H0
Const DT_CENTER = &H1
Const DT_RIGHT = &H2

Private hDcMem As Long
Private hbmpMem As Long
Private hOldBmp As Long
Private bmpWidth As Long
Private bmpHeight As Long
Private UseBrush As Long
Private OldBrush As Long
Private UsePen As Long
Private OldPen As Long
Private rtn As Long
Private FillArea As RECT
Private UseFont As Long
Private mFillColor As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal nColor As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal SourceWidth As Long, ByVal SourceHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function InvertRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText& Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long)
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private hDcMemFace As Long
Private hDcMemLive As Long
Private hdcText As Long
Private hdcText2 As Long
Private m_BkgdColor As OLE_COLOR
Private hDcCtl As Long
Private Colors(2) As Long
Private ColorRect As RECT
Private TextRect As RECT
Private BarsRect As RECT
Private ColorPointer(2, 2) As Integer
Private LastMoved As Integer
Private CurrentColor As Long
Private m_BorderCol As OLE_COLOR
Private Red As Long
Private Green As Long
Private Blue As Long
Private pMouseDown As Boolean
Private m_LabelsFontColor As OLE_COLOR

Event ColorPicked()

Private Sub UserControl_Initialize()
    BkgdColor = vbWhite
    BorderCol = vbBlack
    LabelsFontColor = vbBlack
End Sub
    
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim Index As Integer
    Dim PosX As Integer
    Dim PosY As Integer
    If PtInRect(ColorRect, CLng(X), CLng(Y)) Then
        RaiseEvent ColorPicked
        Exit Sub
    End If
    
    If PtInRect(BarsRect, CLng(X), CLng(Y)) Then
        If Y > 66 Then
            Index = 2
        ElseIf Y > 39 Then
            Index = 1
        Else
            Index = 0
        End If
        pMouseDown = True
        UserControl.SetFocus
        If X < 25 Then X = 25
        If X > 280 Then X = 280
        ColorPointer(Index, 1) = Int(X) - 25
        DrawPicker Index
        DoEvents
    End If
End Sub
    
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim Index As Integer
    If PtInRect(BarsRect, CLng(X), CLng(Y)) Then
        If Y > 66 Then
            Index = 2
        ElseIf Y > 39 Then
            Index = 1
        Else
            Index = 0
        End If
        If pMouseDown Then
            If X < 25 Then X = 25
            If X > 280 Then X = 280
            ColorPointer(Index, 1) = Int(X) - 25
            DrawPicker Index
            DoEvents
        End If
    End If
    
End Sub
    
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If pMouseDown Then
        pMouseDown = False
    End If
    
End Sub
    
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        BkgdColor = .ReadProperty("BkgdColor", vbWhite)
        BorderCol = .ReadProperty("BorderCol", vbBlack)
        LabelsFontColor = .ReadProperty("LabelsFontColor", m_LabelsFontColor)
    End With
    
End Sub
    
Private Sub UserControl_Resize()
    
    UserControl.Height = 1400
    UserControl.Width = 5260
    PrepareControl UserControl.hdc
    'set startup color pointers location here
    ColorPointer(0, 1) = 200
    DrawPicker 0
    ColorPointer(1, 1) = 150
    DrawPicker 1
    ColorPointer(2, 1) = 100
    DrawPicker 2
    UserControl.Refresh
End Sub
    
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BkgdColor", m_BkgdColor, vbWhite
        .WriteProperty "BorderCol", m_BorderCol, vbBlack
        .WriteProperty "LabelsFontColor", m_LabelsFontColor, vbBlack
    End With
End Sub
    
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim X As Single
    Dim NewIndex As Integer
    NewIndex = LastMoved
    If KeyCode = 37 Then    'left arrow
    ColorPointer(LastMoved, 1) = ColorPointer(LastMoved, 1) - 1
    DrawPicker LastMoved
End If
If KeyCode = 39 Then    'right arrow
ColorPointer(LastMoved, 1) = ColorPointer(LastMoved, 1) + 1
DrawPicker LastMoved
End If
End Sub
    
Public Function Color() As Long
    Color = CurrentColor
End Function
    
Private Sub DrawPicker(WhoMoved As Integer)
    
    Dim Point As POINTAPI
    Dim Color As Long
    Dim i As Integer
    Dim Y As Integer
    Dim X As Integer
    LastMoved = WhoMoved
    'Cover Old Arrow from Original Control Face
    Y = 28 + WhoMoved * 27
    X = ColorPointer(WhoMoved, 2) + 19
    rtn = BitBlt(hDcMemLive, X, Y, 16, 8, hDcMemFace, X, Y, vbSrcCopy)
    ColorPointer(WhoMoved, 2) = ColorPointer(WhoMoved, 1) 'index 2 contain previous info
    'Draw New Arrow
    DrawNewArrow WhoMoved
    Red = ColorPointer(0, 1)
    Green = ColorPointer(1, 1)
    Blue = ColorPointer(2, 1)
    'Draw color bars for the 2 colors that did not move
    For i = 0 To 255
        X = i + 26
        If WhoMoved <> 0 Then
            Color = RGB(i, Green, Blue)
            UsePen = GetPen(PS_SOLID, 1, Color)
            Call MoveToEx(hDcMemLive, X, 12, Point)
            Call LineTo(hDcMemLive, X, 28)
            DeletePen
        End If
        If WhoMoved <> 1 Then
            Color = RGB(Red, i, Blue)
            UsePen = GetPen(PS_SOLID, 1, Color)
            Call MoveToEx(hDcMemLive, X, 39, Point)
            Call LineTo(hDcMemLive, X, 55)
            DeletePen
        End If
        If WhoMoved <> 2 Then
            Color = RGB(Red, Green, i)
            UsePen = GetPen(PS_SOLID, 1, Color)
            Call MoveToEx(hDcMemLive, X, 66, Point)
            Call LineTo(hDcMemLive, X, 82)
            DeletePen
        End If
    Next
    UsePen = GetPen(PS_SOLID, 1, BorderCol)
    'Red colorbar border
    Call MoveToEx(hDcMemLive, 26, 12, Point) 'left line
    Call LineTo(hDcMemLive, 26, 27)
    
    Call MoveToEx(hDcMemLive, 281, 12, Point) 'right line
    Call LineTo(hDcMemLive, 281, 28)
    
    Call MoveToEx(hDcMemLive, 26, 12, Point) 'top line
    Call LineTo(hDcMemLive, 281, 12)
    
    Call MoveToEx(hDcMemLive, 26, 27, Point) 'bottom line
    Call LineTo(hDcMemLive, 281, 27)
    
    'Green colorbar border
    Call MoveToEx(hDcMemLive, 26, 39, Point)
    Call LineTo(hDcMemLive, 26, 54)
    
    Call MoveToEx(hDcMemLive, 281, 39, Point)
    Call LineTo(hDcMemLive, 281, 55)
    
    Call MoveToEx(hDcMemLive, 26, 39, Point)
    Call LineTo(hDcMemLive, 281, 39)
    
    Call MoveToEx(hDcMemLive, 26, 54, Point)
    Call LineTo(hDcMemLive, 281, 54)
    
    'Blue colorbar border
    Call MoveToEx(hDcMemLive, 26, 66, Point)
    Call LineTo(hDcMemLive, 26, 81)
    
    Call MoveToEx(hDcMemLive, 281, 66, Point)
    Call LineTo(hDcMemLive, 281, 82)
    
    Call MoveToEx(hDcMemLive, 26, 66, Point)
    Call LineTo(hDcMemLive, 281, 66)
    
    Call MoveToEx(hDcMemLive, 26, 81, Point)
    Call LineTo(hDcMemLive, 281, 81)
    DeletePen
    CurrentColor = RGB(Red, Green, Blue)
    DoColorBlock CurrentColor
    RaiseEvent ColorPicked
    If WhoMoved = 0 Then
        DoText Red, 13
    ElseIf WhoMoved = 1 Then
        DoText Green, 40
    Else
        DoText Blue, 67
    End If
    rtn = BitBlt(UserControl.hdc, 0, 0, 351, 95, hDcMemLive, 0, 0, SRCCOPY)
    UserControl.Refresh
    DoEvents
    
End Sub
    
Private Sub DoColorBlock(Color As Long)
    
    UseBrush = GetBrush(Color)
    UsePen = GetPen(PS_SOLID, 1, BorderCol)
    rtn = Rectangle(hDcMemLive, 316, 10, 346, 82)
    DeleteBrush
    DeletePen
End Sub
    
Private Sub DrawNewArrow(WhoMoved As Integer)
    
    Dim Y As Integer
    Dim Color As Long
    Dim PointArray(2) As POINTAPI
    Y = 28 + WhoMoved * 27
    Colors(WhoMoved) = ColorPointer(WhoMoved, 1)
    If WhoMoved = 0 Then
        Color = RGB(255, 0, 0)
    End If
    If WhoMoved = 1 Then
        Color = RGB(0, 255, 0)
    End If
    If WhoMoved = 2 Then
        Color = &HFF8080
    End If
    Colors(WhoMoved) = 0
    UseBrush = GetBrush(Color)
    UsePen = GetPen(PS_SOLID, 1, BorderCol)
    PointArray(0).X = ColorPointer(WhoMoved, 1) + 26
    PointArray(0).Y = Y
    PointArray(1).X = ColorPointer(WhoMoved, 1) + 33
    PointArray(1).Y = Y + 7
    PointArray(2).X = ColorPointer(WhoMoved, 1) + 19
    PointArray(2).Y = Y + 7
    rtn = Polygon(hDcMemLive, PointArray(0), 3)
    DeleteBrush
    DeletePen
    
End Sub
    
Private Sub DoText(Color As Long, Y As Long)
    
    Dim sColor As String
    rtn = BitBlt(hdcText, 0, 0, 22, 12, hdcText, 0, 0, WHITENESS)
    rtn = BitBlt(hDcMemLive, 289, Y + 1, 22, 12, hDcMemFace, 289, Y + 1, vbSrcCopy)
    sColor = Format(Color, "000")
    rtn = DrawText(hdcText, sColor, Len(sColor), TextRect, DT_CENTER)
    rtn = BitBlt(hDcMemLive, 289, Y + 1, 22, 12, hdcText, 0, 0, vbSrcCopy)
    
End Sub
    
Private Sub PrepareControl(hdc As Long)
    Dim i As Integer
    Dim Point As POINTAPI
    hDcCtl = hdc                'hDc from UserControl
    ColorRect.Left = 315        'The Selected Color on the Control Face
    ColorRect.Top = 14
    ColorRect.Right = 344
    ColorRect.Bottom = 43
    TextRect.Left = 0           'Used to Print Color Values
    TextRect.Top = 0
    TextRect.Right = 20
    TextRect.Bottom = 16
    BarsRect.Left = 15          'The area containing the three color bars on the Control Face
    BarsRect.Top = 10
    BarsRect.Right = 295
    BarsRect.Bottom = 99
    
    'Create Memory BitMaps for Text
    hdcText = Create(21, 12)
    SetFont hDcCtl
    hdcText2 = Create(21, 12)
    'Create a Memory BitMap for the Control Face
    hDcMemFace = Create(351, 95)
    Fill BkgdColor, 0, 0, 351, 95
    'Create a second Memory Bitmap for the Control Face
    hDcMemLive = Create(351, 95)
    Fill BkgdColor, 0, 0, 351, 95
    
    SetTextColor hDcMemFace, LabelsFontColor
    SetTextColor hdcText, LabelsFontColor
    SetTextColor hdcText2, LabelsFontColor
    
    For i = 0 To 2
        If i = 0 Then
            rtn = DrawText(hDcMemFace, "R", 1, TextRect, DT_CENTER)
            rtn = BitBlt(hDcMemLive, 8, 14, 34, 12, hDcMemFace, 0, 2, vbSrcCopy)
        End If
        If i = 1 Then
            rtn = DrawText(hDcMemFace, "G", 1, TextRect, DT_CENTER)
            rtn = BitBlt(hDcMemLive, 8, 41, 34, 12, hDcMemFace, 0, 2, vbSrcCopy)
        End If
        If i = 2 Then
            rtn = DrawText(hDcMemFace, "B", 1, TextRect, DT_CENTER)
            rtn = BitBlt(hDcMemLive, 8, 68, 34, 12, hDcMemFace, 0, 2, vbSrcCopy)
        End If
        DeleteBrush
    Next i
    UsePen = GetPen(PS_SOLID, 1, BorderCol)
    'UC border
    Call MoveToEx(hDcMemLive, 0, 0, Point)
    Call LineTo(hDcMemLive, 0, 93)
    
    Call MoveToEx(hDcMemLive, 350, 0, Point)
    Call LineTo(hDcMemLive, 350, 93)
    
    Call MoveToEx(hDcMemLive, 0, 0, Point)
    Call LineTo(hDcMemLive, 350, 0)
    
    Call MoveToEx(hDcMemLive, 0, 92, Point)
    Call LineTo(hDcMemLive, 350, 92)
    
    'red value box
    Call MoveToEx(hDcMemLive, 287, 12, Point)
    Call LineTo(hDcMemLive, 287, 27)
    
    Call MoveToEx(hDcMemLive, 311, 12, Point)
    Call LineTo(hDcMemLive, 311, 27)
    
    Call MoveToEx(hDcMemLive, 287, 12, Point)
    Call LineTo(hDcMemLive, 311, 12)
    
    Call MoveToEx(hDcMemLive, 287, 27, Point)
    Call LineTo(hDcMemLive, 312, 27)
    
    'green value box
    Call MoveToEx(hDcMemLive, 287, 39, Point)
    Call LineTo(hDcMemLive, 287, 54)
    
    Call MoveToEx(hDcMemLive, 311, 39, Point)
    Call LineTo(hDcMemLive, 311, 55)
    
    Call MoveToEx(hDcMemLive, 287, 39, Point)
    Call LineTo(hDcMemLive, 311, 39)
    
    Call MoveToEx(hDcMemLive, 287, 54, Point)
    Call LineTo(hDcMemLive, 311, 54)
    
    'blue value box
    Call MoveToEx(hDcMemLive, 287, 66, Point)
    Call LineTo(hDcMemLive, 287, 81)
    
    Call MoveToEx(hDcMemLive, 311, 66, Point)
    Call LineTo(hDcMemLive, 311, 82)
    
    Call MoveToEx(hDcMemLive, 287, 66, Point)
    Call LineTo(hDcMemLive, 311, 66)
    
    Call MoveToEx(hDcMemLive, 287, 81, Point)
    Call LineTo(hDcMemLive, 311, 81)
    'R box
    Call MoveToEx(hDcMemLive, 11, 13, Point)
    Call LineTo(hDcMemLive, 11, 27)
    
    Call MoveToEx(hDcMemLive, 11, 12, Point)
    Call LineTo(hDcMemLive, 24, 12)
    
    Call MoveToEx(hDcMemLive, 24, 12, Point)
    Call LineTo(hDcMemLive, 24, 27)
    
    Call MoveToEx(hDcMemLive, 11, 27, Point)
    Call LineTo(hDcMemLive, 25, 27)
    'G box
    Call MoveToEx(hDcMemLive, 11, 39, Point)
    Call LineTo(hDcMemLive, 11, 54)
    
    Call MoveToEx(hDcMemLive, 11, 39, Point)
    Call LineTo(hDcMemLive, 24, 39)
    
    Call MoveToEx(hDcMemLive, 24, 39, Point)
    Call LineTo(hDcMemLive, 24, 54)
    
    Call MoveToEx(hDcMemLive, 11, 54, Point)
    Call LineTo(hDcMemLive, 25, 54)
    'B box
    Call MoveToEx(hDcMemLive, 11, 67, Point)   'left
    Call LineTo(hDcMemLive, 11, 81)
    
    Call MoveToEx(hDcMemLive, 11, 66, Point)  'top
    Call LineTo(hDcMemLive, 24, 66)
    
    Call MoveToEx(hDcMemLive, 24, 66, Point)  'right
    Call LineTo(hDcMemLive, 24, 81)
    
    Call MoveToEx(hDcMemLive, 11, 81, Point)  'bottom
    Call LineTo(hDcMemLive, 25, 81)
    DeletePen
End Sub
    
Public Function getRedValue() As Long
    getRedValue = Red
End Function
Public Function getInvRedValue() As String
    getInvRedValue = 255 - Red
End Function
Public Function getGreenValue() As Long
    getGreenValue = Green
End Function
    
Public Function getInvGreenValue() As String
    getInvGreenValue = 255 - Green
End Function
    
Public Function getBlueValue() As Long
    getBlueValue = Blue
End Function
    
Public Function getInvBlueValue() As String
    getInvBlueValue = 255 - Blue
End Function
    
Public Sub setRedValue(value As Long)
    ColorPointer(0, 1) = value
    DrawPicker 0
End Sub
    
Public Sub setGreenValue(value As Long)
    ColorPointer(1, 1) = value
    DrawPicker 1
End Sub
    
Public Sub setBlueValue(value As Long)
    ColorPointer(2, 1) = value
    DrawPicker 2
End Sub
    
Private Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
    
Private Property Get hdc() As Long
    hdc = UserControl.hdc
End Property
    
Public Property Get BkgdColor() As OLE_COLOR
    BkgdColor = m_BkgdColor
End Property
    
Public Property Let BkgdColor(ByVal NewBkgdColor As OLE_COLOR)
    m_BkgdColor = NewBkgdColor
    PropertyChanged "BkgdColor"
    UserControl_Resize
End Property
    
Public Property Get BorderCol() As OLE_COLOR
    BorderCol = m_BorderCol
End Property
    
Public Property Let BorderCol(ByVal NewBorderCol As OLE_COLOR)
    m_BorderCol = NewBorderCol
    PropertyChanged "BorderCol"
    UserControl_Resize
End Property
    
Public Property Get LabelsFontColor() As OLE_COLOR
    LabelsFontColor = m_LabelsFontColor
End Property
    
Public Property Let LabelsFontColor(ByVal NewLabelsFontColor As OLE_COLOR)
    m_LabelsFontColor = NewLabelsFontColor
    PropertyChanged "LabelsFontColor"
    UserControl_Resize
End Property
    
Public Function Create(WidthIn As Long, HeightIn As Long) As Long
    
    Dim hWndScn As Long
    Dim hDCScn As Long
    bmpWidth = WidthIn
    bmpHeight = HeightIn
    hWndScn = GetDesktopWindow()
    hDCScn = GetDC(hWndScn)
    hDcMem = CreateCompatibleDC(hDCScn)
    hbmpMem = CreateCompatibleBitmap(hDCScn, WidthIn, HeightIn)
    hOldBmp = SelectObject(hDcMem, hbmpMem)
    rtn = BitBlt(hDcMem, 0, 0, WidthIn, HeightIn, hDCScn, 0, 0, WHITENESS)
    mFillColor = vbWhite
    Call ReleaseDC(hWndScn, hDCScn)
    Create = hDcMem
    
End Function
    
Public Function GetPen(ByVal PenType As Long, ByVal PenWidth As Long, ByVal Color As Long) As Long
    
    UsePen = CreatePen(PenType, PenWidth, Color)
    OldPen = SelectObject(hDcMem, UsePen)
    GetPen = UsePen
    
End Function
    
Public Function DeletePen()
    
    rtn = SelectObject(hDcMem, OldPen)
    rtn = DeleteObject(UsePen)
    
End Function
    
Public Function DeleteBrush()
    
    rtn = SelectObject(hDcMem, OldBrush)
    rtn = DeleteObject(UseBrush)
    
End Function
    
Public Function GetBrush(ByVal Color As Long) As Long
    
    UseBrush = CreateSolidBrush(Color)
    OldBrush = SelectObject(hDcMem, UseBrush)
    GetBrush = UseBrush
    
End Function
    
Public Function Copy(Dest As Object)
    
    rtn = BitBlt(Dest.hdc, 0, 0, bmpWidth, bmpHeight, hDcMem, 0, 0, SRCCOPY)
    
End Function
    
Public Function Fill(Color As Long, Optional StartX As Long = 0, Optional StartY As Long = 0, _
    Optional FillWidth As Long = 0, Optional FillHeight As Long = 0)
    
    Dim MyBrush As Long
    MyBrush = GetBrush(Color)
    FillArea.Left = StartX
    FillArea.Top = StartY
    If FillWidth = 0 Then
        FillArea.Right = bmpWidth - StartX
    Else
        FillArea.Right = FillWidth + StartX
    End If
    If FillHeight = 0 Then
        FillArea.Bottom = bmpHeight - StartY
    Else
        FillArea.Bottom = FillHeight + StartY
    End If
    If FillArea.Right > bmpWidth Then FillArea.Right = bmpWidth
    If FillArea.Bottom > bmpHeight Then FillArea.Bottom = bmpHeight
    rtn = FillRect(hDcMem, FillArea, UseBrush)
    DeleteBrush
    mFillColor = Color
    
End Function
    
Public Function SetFont(UsehDc As Long)
    
    UseFont = SelectObject(UsehDc, GetStockObject(SYSTEM_FONT))     'get handle to font
    rtn = SelectObject(UsehDc, UseFont)                             'set UsehDc font back
    rtn = SelectObject(hDcMem, UseFont)                             'set font of MemBmp
    
End Function
    
    
