VERSION 5.00
Begin VB.UserControl ucKF_Slider 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   ScaleHeight     =   930
   ScaleWidth      =   4440
   ToolboxBitmap   =   "ucKF_Slider.ctx":0000
   Begin VB.PictureBox picSlider 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   15
      ScaleHeight     =   345
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   0
      Width           =   390
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         Visible         =   0   'False
         X1              =   180
         X2              =   180
         Y1              =   0
         Y2              =   315
      End
      Begin VB.Line Line3 
         Visible         =   0   'False
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   345
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   390
      End
   End
   Begin VB.Line Line1 
      X1              =   540
      X2              =   540
      Y1              =   0
      Y2              =   435
   End
   Begin VB.Line Line2 
      X1              =   1020
      X2              =   1020
      Y1              =   0
      Y2              =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   210
      Left            =   3855
      TabIndex        =   1
      Top             =   315
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Line linGroove 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   0
      X2              =   4380
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Line linGroove 
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   4380
      Y1              =   180
      Y2              =   180
   End
End
Attribute VB_Name = "ucKF_Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eShape
   RECT = 0
   Squar = 1
   Oval = 2
   Circ = 3
   RdRect = 4
   RdSquar = 5
End Enum

'Property storage
Private lngMin              As Long         'Minimum value range
Private lngMax              As Long         'Maximum value range
Private lngValue            As Long         'Current Value
Private lngSliderWidth      As Long
Private mCenterTick         As Boolean
Private mShowValue          As Boolean
Private mSliderShape        As eShape
Private mSliderBackColor    As OLE_COLOR
Private mLineColor          As OLE_COLOR
Dim lngwidthtemp            As Long
'Event Stubs
Event Changed()
Event MouseUp()

Private Sub picSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseUp
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    UserControl.BackColor = Ambient.BackColor
    picSlider.BackColor = Ambient.BackColor
    SliderBackColor = Ambient.BackColor
End Sub

Private Sub UserControl_InitProperties()
    lngMin = 0
    lngMax = 100
    lngValue = 0
    lngSliderWidth = 200
    UserControl.BackColor = Ambient.BackColor
    picSlider.BackColor = Ambient.BackColor
    Shape1.FillColor = SliderColor
    SliderBackColor = Ambient.BackColor
    SliderShape = 0
    PositionSlider
    LineColor = vbBlack
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        lngMin = .ReadProperty("Min", 0)
        lngMax = .ReadProperty("Max", 100)
        lngValue = .ReadProperty("Value", 50)
        lngSliderWidth = .ReadProperty("SliderWidth", 200)
        SliderColor = .ReadProperty("SliderColor", vb3DFace)
        CenterTick = .ReadProperty("CenterTick", False)
        ShowValue = .ReadProperty("ShowValue", True)
        SliderShape = .ReadProperty("SliderShape", 0)
        SliderBackColor = .ReadProperty("SliderBackColor", Ambient.BackColor)
        LineColor = .ReadProperty("LineColor", vbBlack)
        
    End With
    PositionSlider
End Sub

Private Sub UserControl_Resize()
Dim lngWidth                As Long             'New control width
Dim lngHeight               As Long             'New control height
Dim tc                      As Integer
Dim intIndex                As Integer
    With UserControl
        .Cls
        If ShowValue = True Then
           lngWidth = .Width - 400 - Screen.TwipsPerPixelX
        Else
           lngWidth = .Width - Screen.TwipsPerPixelX
        End If
        lngwidthtemp = lngWidth
        lngHeight = .Height - Screen.TwipsPerPixelY
        UserControl.BackColor = mSliderBackColor
        picSlider.BackColor = mSliderBackColor
        Shape1.FillColor = SliderColor
            For intIndex = 0 To 1
                linGroove(intIndex).X1 = 0
                linGroove(intIndex).X2 = lngWidth
                linGroove(intIndex).Y1 = lngHeight / 2 + 2
                linGroove(intIndex).Y2 = lngHeight / 2 - 2
            Next
            
            picSlider.Top = 0
            picSlider.Height = UserControl.Height
            picSlider.Width = lngSliderWidth
            
    End With
    
    PositionSlider
    'lines 1 & 2 are the two end markers
    Line1.X1 = 0
    Line1.Y1 = 0
    Line1.X2 = 0
    Line1.Y2 = UserControl.Height
    If ShowValue = True Then
       Label1.Left = UserControl.Width - 405
       Line2.X1 = UserControl.Width - 415
       Line2.Y2 = 0
       Line2.X2 = UserControl.Width - 415
       Label1.Visible = True
       Label1.Top = UserControl.Height / 2 - Label1.Height / 2
    Else
       Label1.Left = UserControl.Width - 5
       Line2.X1 = UserControl.Width - 15
       Line2.Y2 = 0
       Line2.X2 = UserControl.Width - 15
       Label1.Visible = False
    End If
    Line2.Y2 = UserControl.Height
    'lines 3 & 4 are in picSliderbox. when value is zero ,3 is visible. when value is max 4 is visible
    Line3.X1 = 0
    Line3.Y1 = 0
    Line3.X2 = 0
    Line3.Y2 = UserControl.Height
    
    Line4.X1 = SliderWidth - 15
    Line4.Y1 = 0
    Line4.X2 = SliderWidth - 15
    Line4.Y2 = UserControl.Height
    
    If lngValue = lngMin Then
       Line3.Visible = True
    Else
       Line3.Visible = False
    End If
                
    If lngValue = lngMax Then
       Line4.Visible = True
    Else
       Line4.Visible = False
    End If
     'center tick mark
    If CenterTick = True Then UserControl.Line (lngWidth / 2, 25)-(lngWidth / 2, UserControl.Height - 25), LineColor
End Sub

Private Sub UserControl_Show()
   ' UserControl.BackColor = Ambient.BackColor
   ' picSlider.BackColor = Ambient.BackColor
  '  SliderBackColor = Ambient.BackColor
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Min", lngMin, 0
        .WriteProperty "Max", lngMax, 100
        .WriteProperty "Value", lngValue, 50
        .WriteProperty "SliderWidth", lngSliderWidth, 200
        .WriteProperty "SliderColor", Shape1.FillColor, vb3DFace
        .WriteProperty "CenterTick", mCenterTick, False
        .WriteProperty "ShowValue", mShowValue, True
        .WriteProperty "SliderShape", mSliderShape, 0
        .WriteProperty "SliderBackColor", mSliderBackColor, Ambient.BackColor
        .WriteProperty "LineColor", mLineColor, vbBlack
    End With
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngPos                  As Long         'New position of slider
Dim sglScale                As Single       'Calculated scale of slider

    With picSlider
            'Caluclate new position and round to nearest pixel
            lngPos = ((X - lngSliderWidth / 2) \ 15) * 15
            'Constrain to control
            If lngPos < 0 Then lngPos = 0
            If ShowValue = True Then
               If lngPos > UserControl.Width - lngSliderWidth - 400 Then
                'Attempted to move past end
                   lngPos = UserControl.Width - lngSliderWidth - 400
                End If
            Else
               If lngPos > UserControl.Width - lngSliderWidth Then
                   lngPos = UserControl.Width - lngSliderWidth
               End If
            End If
            .Left = lngPos
            'Calculate value based on new position
            If ShowValue = True Then
               sglScale = (UserControl.Width - .Width - 401) / (lngMax - lngMin)
            Else
               sglScale = (UserControl.Width - .Width) / (lngMax - lngMin)
            End If
            lngValue = (lngPos / sglScale) + lngMin
            Label1.Caption = lngValue
            If lngValue = lngMin Then
              Line3.Visible = True
            Else
              Line3.Visible = False
            End If
                
            If lngValue = lngMax Then
              Line4.Visible = True
            Else
              Line4.Visible = False
            End If
            RaiseEvent Changed
    End With
End Sub

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngPos                  As Long         'New position of slider
Dim sglScale                As Single       'Calculated scale of slider

    If Button = vbLeftButton Then
        With picSlider
                'calulate new position of slider and round to nearest pixel
                lngPos = ((.Left + X - lngSliderWidth / 2) \ 15) * 15
                    
                'Constrain to control
                If lngPos < 0 Then lngPos = 0
                    'Attempted to move slider past star
                If ShowValue = True Then
                   If lngPos > UserControl.Width - lngSliderWidth - 400 Then
                       lngPos = UserControl.Width - lngSliderWidth - 400
                    End If
                Else
                   If lngPos > UserControl.Width - lngSliderWidth Then
                      lngPos = UserControl.Width - lngSliderWidth
                   End If
                End If
                
                'Move slider
                .Left = lngPos
                
                'Re-calculate value based on new position
                If ShowValue = True Then
                   sglScale = ((UserControl.Width - 401) - lngSliderWidth) / (lngMax - lngMin)
                Else
                   sglScale = ((UserControl.Width) - lngSliderWidth) / (lngMax - lngMin)
                End If
                lngValue = (lngPos / sglScale) + lngMin
                Label1.Caption = lngValue
                If lngValue = lngMin Then
                   Line3.Visible = True
                Else
                   Line3.Visible = False
                End If
                
                If lngValue = lngMax Then
                   Line4.Visible = True
                Else
                   Line4.Visible = False
                End If
                RaiseEvent Changed
        End With
    End If
End Sub

Private Sub picSlider_Resize()
    picSlider.Cls
    picSlider.Height = UserControl.Height
    Shape1.Height = picSlider.Height
    Shape1.Width = picSlider.Width
End Sub

Private Function PositionSlider()
Dim sglScale                As Single       'Calculated scale of slider

    With picSlider
    
        If lngMax - lngMin <> 0 Then
            'Calculate new position
            If ShowValue = True Then
               sglScale = (UserControl.Width - lngSliderWidth - 400) / (lngMax - lngMin)
            Else
               sglScale = (UserControl.Width - lngSliderWidth) / (lngMax - lngMin)
            End If
            .Left = (lngValue - lngMin) * sglScale
        End If
    End With
End Function

Public Property Get SliderBackColor() As OLE_COLOR
    SliderBackColor = mSliderBackColor
End Property

Public Property Let SliderBackColor(NewValue As OLE_COLOR)
    mSliderBackColor = NewValue
    PropertyChanged "SliderBackColor"
    UserControl_Resize
End Property

Public Property Get CenterTick() As Boolean
    CenterTick = mCenterTick
End Property

Public Property Let CenterTick(NewValue As Boolean)

    mCenterTick = NewValue
    PropertyChanged "CenterTick"
    UserControl_Resize
End Property

Public Property Get LineColor() As OLE_COLOR
    LineColor = mLineColor
End Property

Public Property Let LineColor(NewValue As OLE_COLOR)
    mLineColor = NewValue
    Line1.Bordercolor = LineColor
    Line2.Bordercolor = LineColor
    Line3.Bordercolor = LineColor
    Line4.Bordercolor = LineColor
    linGroove(0).Bordercolor = LineColor
    linGroove(1).Bordercolor = LineColor
    Label1.ForeColor = LineColor
    Shape1.Bordercolor = LineColor
    If CenterTick = True Then UserControl.Line (lngwidthtemp / 2, 25)-(lngwidthtemp / 2, UserControl.Height - 25), LineColor
    PropertyChanged "LineColor"
End Property

Public Property Get Min() As Long
    Min = lngMin
End Property

Public Property Let Min(NewValue As Long)
    If NewValue <= lngMax Then
        'Min must be less than Max
        lngMin = NewValue
        
        If lngValue < lngMin Then
            'ensure current value still in min-max range
            lngValue = lngMin
            PropertyChanged "Value"
        End If
        
        PositionSlider
        
        PropertyChanged "Min"
    End If
    
End Property

Public Property Get Max() As Long
    Max = lngMax
End Property

Public Property Let Max(NewValue As Long)

    If NewValue > lngMin Then
        'Max must be greater than Min
        lngMax = NewValue
        
        If lngValue > lngMax Then
            'Ensure current value is within new min-max range
            lngValue = lngMax
            PropertyChanged "Value"
        End If
        
        'Re-initialise slider
        PositionSlider
        
        PropertyChanged "Max"
    End If
    
End Property

Public Property Get ShowValue() As Boolean
    ShowValue = mShowValue
End Property

Public Property Let ShowValue(NewValue As Boolean)

    mShowValue = NewValue
    PropertyChanged "ShowValue"
    UserControl_Resize
End Property

Public Property Get SliderWidth() As Long
    SliderWidth = lngSliderWidth
End Property

Public Property Let SliderWidth(NewValue As Long)
    lngSliderWidth = NewValue
    picSlider.Width = lngSliderWidth
    picSlider.Height = UserControl.Height
    PositionSlider
    PropertyChanged "SliderWidth"
    UserControl_Resize
End Property

Public Property Get SliderColor() As OLE_COLOR
    SliderColor = Shape1.FillColor
End Property

Public Property Let SliderColor(NewValue As OLE_COLOR)
    Shape1.FillColor = NewValue
    PropertyChanged "SliderColor"
End Property

Public Property Get SliderShape() As eShape
    SliderShape = mSliderShape
End Property

Public Property Let SliderShape(NewValue As eShape)
    mSliderShape = NewValue
    Select Case mSliderShape
       Case 0: Shape1.Shape = 0
       Case 1: Shape1.Shape = 1
       Case 2: Shape1.Shape = 2
       Case 3: Shape1.Shape = 3
       Case 4: Shape1.Shape = 4
       Case 5: Shape1.Shape = 5
    End Select
    PropertyChanged "SliderShape"
End Property

Public Property Get Value() As Long
    Value = lngValue
End Property

Public Property Let Value(NewValue As Long)

    'Constrain new value to min-max range
    If NewValue < lngMin Then
        NewValue = lngMin
    
    ElseIf NewValue > lngMax Then
        NewValue = lngMax
    End If
    
    lngValue = NewValue
    
    PositionSlider
    
    PropertyChanged "Value"
    RaiseEvent Changed
    Label1.Caption = NewValue
End Property

