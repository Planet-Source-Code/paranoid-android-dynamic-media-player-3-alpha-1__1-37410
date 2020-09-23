VERSION 5.00
Begin VB.UserControl SliderControl 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E37200&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   ScaleHeight     =   9
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   126
   Begin VB.Shape myShape 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H000080FF&
      BorderStyle     =   0  'Transparent
      BorderWidth     =   2
      FillColor       =   &H00A45200&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "SliderControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************
'       SimpleColorSlider with Float Values        *
'          Written by Andrew Stopakevich           *
'            Modified by Kenneth Hedman            *
'***************************************************

'Default Property Values:
Option Explicit
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 0
'Property Variables:
Dim m_Min As Double
Dim m_Max As Double
Dim m_Enabled As Boolean
Dim m_Vertical As Boolean
Dim m_Value As Double
Dim m_roundto As Double
Dim m_Invert As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_IBackColor As OLE_COLOR
Dim m_SliderColor As OLE_COLOR
Dim m_ISliderColor As OLE_COLOR
'Event Declarations:
Event MouseMove()
Event MouseDown()
Event MouseUp()
Event Click()

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled

End Property

Public Property Let Enabled(New_Value As Boolean)

    m_Enabled = New_Value
    RefreshMe
    PropertyChanged "Enabled"

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor

End Property

Public Property Let BackColor(New_Value As OLE_COLOR)

    m_BackColor = New_Value
    RefreshMe
    PropertyChanged "BackColor"

End Property

Public Property Get SliderColor() As OLE_COLOR

    SliderColor = m_SliderColor

End Property

Public Property Let SliderColor(New_Value As OLE_COLOR)

    m_SliderColor = New_Value
    RefreshMe
    PropertyChanged "SliderColor"

End Property

Public Property Get InvertedBackColor() As OLE_COLOR

    InvertedBackColor = m_IBackColor

End Property

Public Property Let InvertedBackColor(New_Value As OLE_COLOR)

    m_IBackColor = New_Value
    RefreshMe
    PropertyChanged "InvertedBackColor"

End Property

Public Property Get InvertedSliderColor() As OLE_COLOR

    InvertedSliderColor = m_ISliderColor

End Property

Public Property Let InvertedSliderColor(New_Value As OLE_COLOR)

    m_ISliderColor = New_Value
    RefreshMe
    PropertyChanged "InvertedSliderColor"

End Property

Public Property Get Invert() As Boolean

    Invert = m_Invert

End Property

Public Property Let Invert(New_Value As Boolean)

    m_Invert = New_Value
    RefreshMe
    PropertyChanged "Invert"

End Property

Public Property Get Vertical() As Boolean

    Vertical = m_Vertical

End Property

Public Property Let Vertical(New_Value As Boolean)

    m_Vertical = New_Value
    RefreshMe
    PropertyChanged "Vertical"

End Property

Public Property Get Max() As Double

    Max = m_Max

End Property

Public Property Let Max(ByVal New_Max As Double)

    m_Max = New_Max
    PropertyChanged "Max"
    RefreshMe

End Property

'Min, Max
Public Property Get Min() As Double

    Min = m_Min

End Property

Public Property Let Min(ByVal New_Min As Double)

    m_Min = New_Min
    PropertyChanged "Min"
    RefreshMe

End Property

'Positions
Public Sub RefreshMe()

    If m_Max = m_Min Then m_Max = m_Max + 1
    If m_Value < m_Min Then m_Value = m_Min
    If m_Value > m_Max Then m_Value = m_Max
    
    If m_Invert = True Then
        UserControl.BackColor = m_IBackColor
        If m_Value = 0 And m_Vertical = True Then
            myShape.FillColor = m_ISliderColor
        Else
            myShape.FillColor = m_ISliderColor
        End If
    Else
        myShape.FillColor = m_SliderColor
        UserControl.BackColor = m_BackColor
    End If
    
    If m_Enabled = True Then
        UserControl.Enabled = True
      Else
        UserControl.Enabled = False
    End If
    
    If m_Vertical = True Then
        myShape.Height = UserControl.ScaleHeight * (m_Value - m_Min) / (m_Max - m_Min) + 12
        myShape.Width = UserControl.Width
      Else
        myShape.Width = UserControl.ScaleWidth * (m_Value - m_Min) / (m_Max - m_Min) + 12
        myShape.Height = UserControl.Height
    End If

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click

End Sub

'Initalizing the control
Private Sub UserControl_Initialize()

    Call UserControl_Resize

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Value = m_def_Value
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_BackColor = vbBlack
    m_IBackColor = vbWhite
    m_Enabled = True
    m_Vertical = False
    m_Invert = False
    RefreshMe

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> vbLeftButton Then Exit Sub

    If m_Vertical = False Then
        If X = 0 Then Exit Sub ':( Expand Structure or consider reversing Condition
        m_Value = (X * (m_Max - m_Min) + UserControl.ScaleWidth * m_Min) / UserControl.ScaleWidth
        RefreshMe
        myShape.BorderStyle = 1
        
        RaiseEvent MouseDown
    Else
        If Y = 0 Then Exit Sub ':( Expand Structure or consider reversing Condition
        m_Value = (Y * (m_Max - m_Min) + UserControl.ScaleHeight * m_Min) / UserControl.ScaleHeight
        RefreshMe
        myShape.BorderStyle = 1
        
        RaiseEvent MouseDown
    End If
    
End Sub

'Moves the slider
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> vbLeftButton Then Exit Sub

    If m_Vertical = False Then
        If Button = 1 And X > 0 Then
            m_Value = (X * (m_Max - m_Min) + UserControl.ScaleWidth * m_Min) / UserControl.ScaleWidth
            RefreshMe
            RaiseEvent MouseMove
        End If
    Else
        If Button = 1 And Y > 0 Then
            m_Value = (Y * (m_Max - m_Min) + UserControl.ScaleHeight * m_Min) / UserControl.ScaleHeight
            RefreshMe
            RaiseEvent MouseMove
        End If
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> vbLeftButton Then Exit Sub

    If m_Vertical = False Then
        If X = 0 Then Exit Sub ':( Expand Structure or consider reversing Condition
        m_Value = (X * (m_Max - m_Min) + UserControl.ScaleWidth * m_Min) / UserControl.ScaleWidth
        RefreshMe
        myShape.BorderStyle = 0
        
        RaiseEvent MouseUp
    Else
        If Y = 0 Then Exit Sub ':( Expand Structure or consider reversing Condition
        m_Value = (Y * (m_Max - m_Min) + UserControl.ScaleHeight * m_Min) / UserControl.ScaleHeight
        RefreshMe
        myShape.BorderStyle = 0
        
        RaiseEvent MouseUp
    End If
    
End Sub


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_Vertical = PropBag.ReadProperty("Vertical", False)
    m_Invert = PropBag.ReadProperty("Invert", False)
    m_BackColor = PropBag.ReadProperty("BackColor", vbBlack)
    m_IBackColor = PropBag.ReadProperty("InvertedBackColor", vbWhite)
    m_SliderColor = PropBag.ReadProperty("SliderBackColor", vbBlack)
    m_ISliderColor = PropBag.ReadProperty("InvertedSliderColor", vbWhite)

    RefreshMe

End Sub

'Resize sub
Private Sub UserControl_Resize()

    If m_Vertical = True Then
        myShape.Height = UserControl.Height + 10
        myShape.left = -10
        myShape.top = -10
        RefreshMe
    Else
        myShape.Width = UserControl.Width + 10
        myShape.left = -10
        myShape.top = -10
        RefreshMe
    End If
    
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Vertical", m_Vertical, False)
    Call PropBag.WriteProperty("Invert", m_Invert, False)
    Call PropBag.WriteProperty("BackColor", m_BackColor, vbBlack)
    Call PropBag.WriteProperty("InvertedBackColor", m_IBackColor, vbWhite)
    Call PropBag.WriteProperty("SliderColor", m_SliderColor, vbBlack)
    Call PropBag.WriteProperty("InvertedSliderColor", m_ISliderColor, vbWhite)

End Sub

'Values
Public Property Get Value() As Double

    Value = Round(m_Value, m_roundto)

End Property

Public Property Let Value(ByVal New_Value As Double)

    m_Value = New_Value
    RefreshMe
    PropertyChanged "Value"

End Property
