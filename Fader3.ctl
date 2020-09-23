VERSION 5.00
Begin VB.UserControl Fader 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox DraggerSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   1800
      Picture         =   "Fader3.ctx":0000
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "Fader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This code might be really complicated and hard to understand, im not
'the best of coders. But the idea is that you can use this control
'without having to bother with this code at all :)
Event Scrolling()
    Private Enum RasterOps
        srccopy = &HCC0020
         SRCAND = &H8800C6
         SRCINVERT = &H660046
         SRCPAINT = &HEE0086
         SRCERASE = &H4400328
         WHITENESS = &HFF0062
         BLACKNESS = &H42
    End Enum
     Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As RasterOps _
        ) As Long
Dim mValue As Integer 'what value we are at
Private Const mMax As Integer = 100 'maximum value
Dim Dragging As Boolean 'whether we are dragging or not
Public Property Get Value() As Integer
    Value = mValue
End Property
Public Property Let Value(ByVal NewValue As Integer)
    PropertyChanged "Value"
    mValue = NewValue
    Redraw
End Property
Public Function topboundary() As Integer
   topboundary = 2
End Function
Public Function BottomBoundary() As Integer
   BottomBoundary = UserControl.ScaleHeight - 30
End Function
Public Function OneIncrement() As Double
    OneIncrement = (BottomBoundary - topboundary) / mMax
End Function
Public Sub DrawPicAtValue(tValue As Integer)
    If tValue > mMax Then tValue = mMax 'if new value is too big, make it max size
    If tValue < 0 Then tValue = 0 'if new value is too small, make it 0
    Dim NewY As Integer
    NewY = topboundary + ((mMax - tValue) * OneIncrement) 'calculate new y value of slider
    BitBlt UserControl.hDC, 1, NewY, DraggerSource.Width, DraggerSource.Height, DraggerSource.hDC, 0, 0, srccopy 'draw it
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging = True Then
        Dim realY As Single
        realY = (Y - topboundary) / OneIncrement 'use the Y value of the mouse to calculate the value
        mValue = (mMax - realY) + 15 'set the value
        Redraw 'redraw
        RaiseEvent Scrolling
    End If
End Sub
Public Function FaderHeight() As Integer
    'change fader height here
    FaderHeight = UserControl.ScaleHeight - 1 '(pixels)
End Function
Private Sub Redraw()
    Dim LightColor As Long, DarkColor As Long
    UserControl.Cls
    'change colors here::
    LightColor = vbWhite
    DarkColor = &H8000000C
    'draws the 3D looking border lines (no need to change this code)
    UserControl.Line (0, 0)-(20, 0), DarkColor, BF
    UserControl.Line (0, 0)-(0, FaderHeight - 1), DarkColor, BF
    UserControl.Line (21, 0)-(21, FaderHeight), LightColor, BF
    UserControl.Line (0, FaderHeight)-(21, FaderHeight), LightColor, BF
    'draws the dark rectangle (no need to change this code)
    UserControl.Line (9, 2)-(12, 2), DarkColor, BF
    UserControl.Line (9, 2)-(9, FaderHeight - 3), DarkColor, BF
    UserControl.Line (12, 3)-(12, FaderHeight - 2), LightColor, BF
    UserControl.Line (10, 3)-(11, FaderHeight - 2), vbBlack, BF
    DrawPicAtValue mValue
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = True
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
End Sub
Private Sub UserControl_Resize()
 Redraw
End Sub
