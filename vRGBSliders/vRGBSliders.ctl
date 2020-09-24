VERSION 5.00
Begin VB.UserControl vRGBSliders 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   Palette         =   "vRGBSliders.ctx":0000
   PaletteMode     =   2  'Custom
   ScaleHeight     =   198
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   ToolboxBitmap   =   "vRGBSliders.ctx":0FFA
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   180
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   180
      Width           =   315
   End
   Begin VB.PictureBox picRamp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   60
      MouseIcon       =   "vRGBSliders.ctx":130C
      MousePointer    =   99  'Custom
      Picture         =   "vRGBSliders.ctx":1BD6
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3030
   End
   Begin VB.PictureBox picRGB 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   780
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   120
      Begin VB.Label lblColor 
         Caption         =   "R"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   135
      End
      Begin VB.Label lblColor 
         Caption         =   "G"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   135
      End
      Begin VB.Label lblColor 
         Caption         =   "B"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   720
         Width           =   135
      End
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Index           =   2
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Index           =   1
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Index           =   0
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox picSelRect 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   135
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   10
      Top             =   135
      Width           =   390
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   360
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   315
   End
   Begin VB.Image imgRamp 
      Height          =   225
      Index           =   3
      Left            =   60
      Picture         =   "vRGBSliders.ctx":3F04
      Top             =   2640
      Width           =   2970
   End
   Begin VB.Image imgRamp 
      Height          =   225
      Index           =   2
      Left            =   60
      Picture         =   "vRGBSliders.ctx":4EFE
      Top             =   2340
      Width           =   2970
   End
   Begin VB.Image imgRamp 
      Height          =   225
      Index           =   1
      Left            =   60
      Picture         =   "vRGBSliders.ctx":5EF8
      Top             =   2040
      Width           =   2970
   End
   Begin VB.Image imgRamp 
      Height          =   225
      Index           =   0
      Left            =   60
      Picture         =   "vRGBSliders.ctx":6EF2
      Top             =   1740
      Width           =   2970
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu mnuRamp 
         Caption         =   "RGB Spectrum"
         Index           =   0
      End
      Begin VB.Menu mnuRamp 
         Caption         =   "Greyscale Ramp"
         Index           =   1
      End
      Begin VB.Menu mnuRamp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuRamp 
         Caption         =   "Make Ramp Web Safe"
         Index           =   3
      End
   End
End
Attribute VB_Name = "vRGBSliders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ctlRect As RECT
Dim RedRect As RECT
Dim GreenRect As RECT
Dim BlueRect As RECT
Dim RArrow As RECT
Dim GArrow As RECT
Dim BArrow As RECT
Dim RightRect As RECT
Dim LeftRect As RECT
Dim FromSliders As Boolean
Public Enum BorderStyleEnum
    Flat = 0
    Sunken
    Raised
    Etched
    Bump
End Enum
Public Enum RampStyleEnum
Normal
GreyScale
End Enum
Public Enum FromLocation
cText
cSliders
cGreyScale
End Enum
Public Enum ColorSelectedEnum
Leftcolor = 0
Rightcolor = 1
End Enum
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Const LeftSide = 70
Dim ScrollWidth As Single
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Event Change()
'Default Property Values:
Const m_def_cRampStyle = 0
Const m_def_CurrentColor = 0
Const m_def_BorderStyle = 0
Const m_def_cRedValue = 0
Const m_def_cGreenValue = 0
Const m_def_cBlueValue = 0
'Property Variables:
Dim m_cRampStyle As Integer
Dim m_CurrentColor As OLE_COLOR
Dim m_BorderStyle As Integer
Dim m_cRedValue As Integer
Dim m_cGreenValue As Integer
Dim m_cBlueValue As Integer
Dim SelectedColor As ColorSelectedEnum
Dim SELCOL As Long
Dim Location As FromLocation

Private Sub mnuRamp_Click(Index As Integer)

Select Case Index
Case 0
    Select Case mnuRamp(3).Checked
    Case True
    picRamp.Picture = imgRamp(2).Picture
    m_cRampStyle = 0
    mnuRamp(1).Checked = False
    Case False
    picRamp.Picture = imgRamp(0).Picture
    m_cRampStyle = 0
    mnuRamp(1).Checked = False
    End Select
    mnuRamp(0).Checked = True
Case 1

    Select Case mnuRamp(3).Checked
    Case True
    picRamp.Picture = imgRamp(3).Picture
    m_cRampStyle = 1
    mnuRamp(0).Checked = False
        Case False
    picRamp.Picture = imgRamp(1).Picture
    m_cRampStyle = 1
    mnuRamp(0).Checked = False
    End Select
    mnuRamp(1).Checked = True
    
Case 3

    Select Case m_cRampStyle
    Case 0
        If mnuRamp(3).Checked Then
        picRamp.Picture = imgRamp(0).Picture
        Else
        picRamp.Picture = imgRamp(2).Picture
        End If
        mnuRamp(0).Checked = True
        mnuRamp(1).Checked = False
        m_cRampStyle = 0
    
    Case 1
        If mnuRamp(3).Checked Then
        picRamp.Picture = imgRamp(1).Picture
        Else
        picRamp.Picture = imgRamp(3).Picture
        End If
        mnuRamp(0).Checked = False
        mnuRamp(1).Checked = True
        m_cRampStyle = 1
    End Select
txtColor(0).Text = MakeWebSafe(txtColor(0).Text)
txtColor(1).Text = MakeWebSafe(txtColor(1).Text)
txtColor(2).Text = MakeWebSafe(txtColor(2).Text)
mnuRamp(3).Checked = Not mnuRamp(3).Checked
End Select
    If mnuRamp(3).Checked Then
    mnuRamp(3).Caption = "Show non-Web Safe Ramp"
    Else
    mnuRamp(3).Caption = "Make Ramp Web Safe"
    End If
picRamp.Line (190, 0)-(198, 6), RGB(255, 255, 255), BF
Draw
UpdateRects
End Sub

Private Sub picLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Location = cSliders
SelectedColor = Leftcolor
SELCOL = picLeft.BackColor
    txtColor(0).Text = Red(SELCOL)
    txtColor(1).Text = Green(SELCOL)
    txtColor(2).Text = Blue(SELCOL)
     picSelRect.Move 9, 9, 26, 26
     picRight.ZOrder 1
         Draw
    UpdateRects2
    DrawColorSelectors
     Location = cText
End Sub
Private Sub picRamp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
GetFromRamp Button, Shift, X, Y
End Sub
Private Sub picRamp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
GetFromRamp Button, Shift, X, Y
End Sub
Private Sub picRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Location = cSliders
SelectedColor = Rightcolor
SELCOL = picRight.BackColor
    txtColor(0).Text = Red(SELCOL)
    txtColor(1).Text = Green(SELCOL)
    txtColor(2).Text = Blue(SELCOL)
    picSelRect.Move 21, 21, 26, 26
    picSelRect.ZOrder 0
    picRight.ZOrder 0
    picLeft.ZOrder 0
        Draw
    UpdateRects2
    DrawColorSelectors
    Location = cText
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim MapX As Long
    If X <= LeftSide Then
        X = LeftSide
    End If
    If X >= (UserControl.ScaleWidth - 40) Then
        X = (UserControl.ScaleWidth - 40)
    End If

    If Button = 1 Then
        Location = cSliders
        If mnuRamp(3).Checked Then
       X = X - LeftSide
       Select Case X
       Case 0 To 10
       X = 0
       Case 11 To 30
       X = 20
       Case 31 To 50
       X = 40
       Case 51 To 70
       X = 61
       Case 71 To 90
       X = 81
       Case Is >= 91
       X = 101
       End Select
       
        Select Case Y
        
        Case RedRect.Top To (RedRect.Bottom + 15)
            SetRect RArrow, X + (LeftSide - 5), RedRect.Bottom, X + (LeftSide + 5), RedRect.Bottom + 5
            txtColor(0).Text = MakeWebSafe(X * ScrollWidth)
        Case GreenRect.Top To (GreenRect.Bottom + 15)
            SetRect GArrow, X + (LeftSide - 5), GreenRect.Bottom, X + (LeftSide + 5), GreenRect.Bottom + 5
            txtColor(1).Text = MakeWebSafe(X * ScrollWidth)
        Case BlueRect.Top To (BlueRect.Bottom + 15)
            SetRect BArrow, X + (LeftSide - 5), BlueRect.Bottom, X + (LeftSide + 5), BlueRect.Bottom + 5
            txtColor(2).Text = MakeWebSafe(X * ScrollWidth)
        Case Else
        End Select

        
        Else
        
        Select Case Y
        
        Case RedRect.Top To (RedRect.Bottom + 15)
            SetRect RArrow, X - 5, RedRect.Bottom, X + 5, RedRect.Bottom + 5
            txtColor(0).Text = CInt(((X - LeftSide) * ScrollWidth))
        Case GreenRect.Top To (GreenRect.Bottom + 15)
            SetRect GArrow, X - 5, GreenRect.Bottom, X + 5, GreenRect.Bottom + 5
            txtColor(1).Text = CInt(((X - LeftSide) * ScrollWidth))
        Case BlueRect.Top To (BlueRect.Bottom + 15)
            SetRect BArrow, X - 5, BlueRect.Bottom, X + 5, BlueRect.Bottom + 5
            txtColor(2).Text = CInt(((X - LeftSide) * ScrollWidth))
        Case Else
        End Select
        End If
        
            Select Case SelectedColor
            Case Leftcolor
                picLeft.BackColor = RGB(m_cRedValue, m_cGreenValue, m_cBlueValue)
            Case Rightcolor
                picRight.BackColor = RGB(m_cRedValue, m_cGreenValue, m_cBlueValue)
            End Select
            
    RaiseEvent Change
        Location = cText
    End If
        UpdateRects
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_Resize()
    Width = 3165
    Height = 1545
    Size Width, Height
    SetRect ctlRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    SetRect RedRect, LeftSide, 14, UserControl.ScaleWidth - 38, 22
    SetRect GreenRect, LeftSide, 38, UserControl.ScaleWidth - 38, 46
    SetRect BlueRect, LeftSide, 62, UserControl.ScaleWidth - 38, 70
    SetRect RArrow, RedRect.Left - 5, RedRect.Bottom, RedRect.Left + 5, RedRect.Bottom + 5
    SetRect GArrow, GreenRect.Left - 5, GreenRect.Bottom, GreenRect.Left + 5, GreenRect.Bottom + 5
    SetRect BArrow, BlueRect.Left - 5, BlueRect.Bottom, BlueRect.Left + 5, BlueRect.Bottom + 5
    SetRect RightRect, 0, 0, picRight.Width, picRight.Height
    SetRect LeftRect, 0, 0, picLeft.Width, picLeft.Height
    txtColor(0).Move UserControl.ScaleWidth - 30
    txtColor(1).Move UserControl.ScaleWidth - 30
    txtColor(2).Move UserControl.ScaleWidth - 30
    picRGB.Move LeftSide - 16
    ScrollWidth = (255 / ((UserControl.ScaleWidth - 38 - LeftSide) - 2)) '(255 / ((RedRect.Right - LeftSide)))
    
    Draw
    'MsgBox (ScrollWidth * ((UserControl.ScaleWidth - 38 - LeftSide) - 2)) / 6
    UpdateRects
End Sub
Private Sub Draw()
    DoEvents
    UserControl.Cls

    'Draw Red Gradient
    Gradient RedRect, 255, Val(txtColor(1).Text), Val(txtColor(2).Text), 0, Val(txtColor(1).Text), Val(txtColor(2).Text)
    'Draw Red Edge
    DrawEdge UserControl.hdc, RedRect, BDR_SUNKENINNER, BF_TOPLEFT
    DrawEdge UserControl.hdc, RedRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
    'Draw Green Gradient
    Gradient GreenRect, Val(txtColor(0).Text), 255, Val(txtColor(2).Text), Val(txtColor(0).Text), 0, Val(txtColor(2).Text)
    'Draw Green Edge
    DrawEdge UserControl.hdc, GreenRect, BDR_SUNKENINNER, BF_TOPLEFT
    DrawEdge UserControl.hdc, GreenRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
    'Draw Blue Gradient
    Gradient BlueRect, Val(txtColor(0).Text), Val(txtColor(1).Text), 255, Val(txtColor(0).Text), Val(txtColor(1).Text), 0
    'Draw Blue Edge
    DrawEdge UserControl.hdc, BlueRect, BDR_SUNKENINNER, BF_TOPLEFT
    DrawEdge UserControl.hdc, BlueRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
    'Draw Right & Left Color Selector Borders
    DrawEdge picLeft.hdc, LeftRect, BDR_SUNKENINNER, BF_TOPLEFT
    DrawEdge picLeft.hdc, LeftRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
    
    DrawEdge picRight.hdc, RightRect, BDR_SUNKENINNER, BF_TOPLEFT
    DrawEdge picRight.hdc, RightRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
    
    'Draw Border
    Select Case m_BorderStyle
    Case 0
    
    Case 1
    DrawEdge UserControl.hdc, ctlRect, EDGE_SUNKEN, BF_RECT
    Case 2
    DrawEdge UserControl.hdc, ctlRect, EDGE_RAISED, BF_RECT
    Case 3
    DrawEdge UserControl.hdc, ctlRect, EDGE_ETCHED, BF_RECT
    Case 4
    DrawEdge UserControl.hdc, ctlRect, EDGE_BUMP, BF_RECT
    Case 5
    DrawEdge UserControl.hdc, ctlRect, BDR_SUNKENOUTER, BF_RECT
    End Select
If mnuRamp(3).Checked Then
UserControl.Line (RedRect.Left + 20, RedRect.Top + 3)-(RedRect.Left + 20, RedRect.Bottom - 1)
UserControl.Line (RedRect.Left + 40, RedRect.Top + 3)-(RedRect.Left + 40, RedRect.Bottom - 1)
UserControl.Line (RedRect.Left + 61, RedRect.Top + 3)-(RedRect.Left + 61, RedRect.Bottom - 1)
UserControl.Line (RedRect.Left + 81, RedRect.Top + 3)-(RedRect.Left + 81, RedRect.Bottom - 1)

UserControl.Line (GreenRect.Left + 20, GreenRect.Top + 3)-(GreenRect.Left + 20, GreenRect.Bottom - 1)
UserControl.Line (GreenRect.Left + 40, GreenRect.Top + 3)-(GreenRect.Left + 40, GreenRect.Bottom - 1)
UserControl.Line (GreenRect.Left + 61, GreenRect.Top + 3)-(GreenRect.Left + 61, GreenRect.Bottom - 1)
UserControl.Line (GreenRect.Left + 81, GreenRect.Top + 3)-(GreenRect.Left + 81, GreenRect.Bottom - 1)

UserControl.Line (BlueRect.Left + 20, BlueRect.Top + 3)-(BlueRect.Left + 20, BlueRect.Bottom - 1)
UserControl.Line (BlueRect.Left + 40, BlueRect.Top + 3)-(BlueRect.Left + 40, BlueRect.Bottom - 1)
UserControl.Line (BlueRect.Left + 61, BlueRect.Top + 3)-(BlueRect.Left + 61, BlueRect.Bottom - 1)
UserControl.Line (BlueRect.Left + 81, BlueRect.Top + 3)-(BlueRect.Left + 81, BlueRect.Bottom - 1)


End If
End Sub
Private Sub UserControl_Show()
Draw
UpdateRects
'Rainbow
picRamp.Line (190, 0)-(198, 6), RGB(255, 255, 255), BF
'SavePicture picRamp.Image, App.Path & "\r.bmp"
End Sub
Private Sub UpdateRects2()
Dim MapR As Single
Dim MapG As Single
Dim MapB As Single

MapR = (Val(txtColor(0).Text) / ScrollWidth) + LeftSide - 5
MapG = (Val(txtColor(1).Text) / ScrollWidth) + LeftSide - 5
MapB = (Val(txtColor(2).Text) / ScrollWidth) + LeftSide - 5

UserControl.Line (RedRect.Left - 10, RedRect.Bottom)-(RedRect.Right + 10, RedRect.Bottom + 10), UserControl.BackColor, BF
UserControl.Line (GreenRect.Left - 10, GreenRect.Bottom)-(GreenRect.Right + 10, GreenRect.Bottom + 10), UserControl.BackColor, BF
UserControl.Line (BlueRect.Left - 10, BlueRect.Bottom)-(BlueRect.Right + 10, BlueRect.Bottom + 10), UserControl.BackColor, BF

SetRect RArrow, MapR, RArrow.Top, MapR + 10, RArrow.Bottom
SetRect GArrow, MapG, GArrow.Top, MapG + 10, GArrow.Bottom
SetRect BArrow, MapB, BArrow.Top, MapB + 10, BArrow.Bottom
'Red Cursor
0 Line (RArrow.Left + 5, RArrow.Top)-(RArrow.Left, RArrow.Bottom)
Line -(RArrow.Right, RArrow.Bottom)
Line -(RArrow.Left + 5, RArrow.Top)

'GreenCursor
Line (GArrow.Left + 5, GArrow.Top)-(GArrow.Left, GArrow.Bottom)
Line -(GArrow.Right, GArrow.Bottom)
Line -(GArrow.Left + 5, GArrow.Top)

'Blue Cursor
Line (BArrow.Left + 5, BArrow.Top)-(BArrow.Left, BArrow.Bottom)
Line -(BArrow.Right, BArrow.Bottom)
Line -(BArrow.Left + 5, BArrow.Top)

End Sub
Private Sub UpdateRects()
UserControl.Line (RedRect.Left - 10, RedRect.Bottom)-(RedRect.Right + 10, RedRect.Bottom + 10), UserControl.BackColor, BF
UserControl.Line (GreenRect.Left - 10, GreenRect.Bottom)-(GreenRect.Right + 10, GreenRect.Bottom + 10), UserControl.BackColor, BF
UserControl.Line (BlueRect.Left - 10, BlueRect.Bottom)-(BlueRect.Right + 10, BlueRect.Bottom + 10), UserControl.BackColor, BF
'Red Cursor
Line (RArrow.Left + 5, RArrow.Top)-(RArrow.Left, RArrow.Bottom)
Line -(RArrow.Right, RArrow.Bottom)
Line -(RArrow.Left + 5, RArrow.Top)

'GreenCursor
Line (GArrow.Left + 5, GArrow.Top)-(GArrow.Left, GArrow.Bottom)
Line -(GArrow.Right, GArrow.Bottom)
Line -(GArrow.Left + 5, GArrow.Top)

'Blue Cursor
Line (BArrow.Left + 5, BArrow.Top)-(BArrow.Left, BArrow.Bottom)
Line -(BArrow.Right, BArrow.Bottom)
Line -(BArrow.Left + 5, BArrow.Top)

    DrawEdge picLeft.hdc, LeftRect, BDR_SUNKENINNER, BF_TOPLEFT
    DrawEdge picLeft.hdc, LeftRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
    
    DrawEdge picRight.hdc, RightRect, BDR_SUNKENINNER, BF_TOPLEFT
    DrawEdge picRight.hdc, RightRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
   
End Sub
Private Sub Gradient(cR As RECT, Color1R As Byte, Color1G As Byte, Color1B As Byte, Color2R As Byte, Color2G As Byte, Color2B As Byte)
Dim i As Long
Dim R As Byte
Dim G As Byte
Dim B As Byte
Dim RightFiX As Integer
RightFiX = cR.Right - cR.Left - 2
For i = 0 To RightFiX
  R = ((i * ((Color1R + RightFiX) - Color2R)) / RightFiX + Color2R - i)
  G = ((i * ((Color1G + RightFiX) - Color2G)) / RightFiX + Color2G - i)
  B = ((i * ((Color1B + RightFiX) - Color2B)) / RightFiX + Color2B - i)
UserControl.Line (i + cR.Left, cR.Top)-(i + 1 + cR.Left, cR.Bottom - 1), RGB(R, G, B), BF
Next i
End Sub
Private Sub txtColor_Change(Index As Integer)
    If txtColor(Index).Text = "" Then
        txtColor(Index).Text = 0
        txtColor(Index).SelStart = 0
        txtColor(Index).SelLength = Len(txtColor(Index).Text)
    End If

    If Val(txtColor(Index).Text) < 0 Then txtColor(Index).Text = 0
    If Val(txtColor(Index).Text) > 255 Then txtColor(Index).Text = 255
    m_cRedValue = txtColor(0).Text
    m_cGreenValue = txtColor(1).Text
    m_cBlueValue = txtColor(2).Text
    m_CurrentColor = RGB(m_cRedValue, m_cGreenValue, m_cBlueValue)
    
    
    
    Select Case Location
    Case 0
    Draw
    UpdateRects2
    DrawColorSelectors
    Case 1
    Draw
    Case 2

    'End If
    End Select
    RaiseEvent Change
End Sub
Private Sub txtColor_GotFocus(Index As Integer)
    txtColor(Index).SelStart = 0
    txtColor(Index).SelLength = Len(txtColor(Index).Text)
End Sub
Private Sub txtColor_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Numbers As Integer
        Numbers = KeyAscii
    If ((Numbers < 48 Or Numbers > 57) And Numbers <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Function Red(ByVal Color As Long) As Integer
    Red = Color Mod &H100
End Function
Private Function Green(ByVal Color As Long) As Integer
    Green = (Color \ &H100) Mod &H100
End Function
Private Function Blue(ByVal Color As Long) As Integer
    Blue = (Color \ &H10000) Mod &H100
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get cRedValue() As Integer
Attribute cRedValue.VB_Description = "Returns/sets the red value for an object."
    cRedValue = m_cRedValue
End Property
Public Property Let cRedValue(ByVal New_cRedValue As Integer)
    If New_cRedValue > 255 Then New_cRedValue = 255
    If New_cRedValue < 0 Then New_cRedValue = 0
    m_cRedValue = New_cRedValue
    PropertyChanged "cRedValue"
    txtColor(0).Text = m_cRedValue
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get cGreenValue() As Integer
Attribute cGreenValue.VB_Description = "Returns/sets the green value for an object."
    cGreenValue = m_cGreenValue
End Property
Public Property Let cGreenValue(ByVal New_cGreenValue As Integer)
    If New_cGreenValue > 255 Then New_cGreenValue = 255
    If New_cGreenValue < 0 Then New_cGreenValue = 0
    m_cGreenValue = New_cGreenValue
    PropertyChanged "cGreenValue"
    txtColor(1).Text = m_cGreenValue
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get cBlueValue() As Integer
Attribute cBlueValue.VB_Description = "Returns/sets the blue value for an object."
    cBlueValue = m_cBlueValue
End Property
Public Property Let cBlueValue(ByVal New_cBlueValue As Integer)
    If New_cBlueValue > 255 Then New_cBlueValue = 255
    If New_cBlueValue < 0 Then New_cBlueValue = 0
    m_cBlueValue = New_cBlueValue
    PropertyChanged "cBlueValue"
    txtColor(2).Text = m_cBlueValue
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_cRedValue = m_def_cRedValue
    m_cGreenValue = m_def_cGreenValue
    m_cBlueValue = m_def_cBlueValue
    m_BorderStyle = m_def_BorderStyle
    m_CurrentColor = m_def_CurrentColor
    m_cRampStyle = m_def_cRampStyle
    m_cRampStyle = m_def_cRampStyle
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_cRedValue = PropBag.ReadProperty("cRedValue", m_def_cRedValue)
    m_cGreenValue = PropBag.ReadProperty("cGreenValue", m_def_cGreenValue)
    m_cBlueValue = PropBag.ReadProperty("cBlueValue", m_def_cBlueValue)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_CurrentColor = PropBag.ReadProperty("CurrentColor", m_def_CurrentColor)
    m_cRampStyle = PropBag.ReadProperty("cRampStyle", m_def_cRampStyle)
    picRamp.Picture = imgRamp(m_cRampStyle).Picture
    mnuRamp(m_cRampStyle).Checked = True
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("cRedValue", m_cRedValue, m_def_cRedValue)
    Call PropBag.WriteProperty("cGreenValue", m_cGreenValue, m_def_cGreenValue)
    Call PropBag.WriteProperty("cBlueValue", m_cBlueValue, m_def_cBlueValue)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("CurrentColor", m_CurrentColor, m_def_CurrentColor)
    Call PropBag.WriteProperty("cRampStyle", m_cRampStyle, m_def_cRampStyle)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As BorderStyleEnum
Attribute BorderStyle.VB_Description = "Returns sets The borde rstyle for an object."
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Draw
    UpdateRects
End Property
Private Sub GetFromRamp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CX As Integer
Select Case Button
Case 1
    Select Case m_cRampStyle
        Case 0
        If mnuRamp(3).Checked Then 'See If Using Websafe Ramp
                Location = cGreyScale
                SELCOL = (GetPixel(picRamp.hdc, X, Y))
                txtColor(0).Text = MakeWebSafe(Red(SELCOL))
                txtColor(1).Text = MakeWebSafe(Green(SELCOL))
                txtColor(2).Text = MakeWebSafe(Blue(SELCOL))
                Draw
                UpdateRects2
                Location = cText
        Else
                Location = cText
                SELCOL = (GetPixel(picRamp.hdc, X, Y))
                txtColor(0).Text = Red(SELCOL)
                txtColor(1).Text = Green(SELCOL)
                txtColor(2).Text = Blue(SELCOL)
                Location = cText
        End If
        Case 1
                Location = cGreyScale
                SELCOL = (GetPixel(picRamp.hdc, X, Y))
                CX = MakeGray(Red(SELCOL), Green(SELCOL), Blue(SELCOL))
                txtColor(0).Text = CX
                txtColor(1).Text = CX
                txtColor(2).Text = CX
                Draw
                UpdateRects2
                Location = cText
                
                
    End Select
Case 2
    PopupMenu mnuPopup, , X, picRamp.Top
End Select
    DrawColorSelectors
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CurrentColor() As OLE_COLOR
Attribute CurrentColor.VB_Description = "Returns the current selected color for an object."
    CurrentColor = m_CurrentColor
End Property
Private Sub DrawColorSelectors()
    Select Case SelectedColor
    Case Leftcolor
        picLeft.BackColor = RGB(m_cRedValue, m_cGreenValue, m_cBlueValue)
    Case Rightcolor
        picRight.BackColor = RGB(m_cRedValue, m_cGreenValue, m_cBlueValue)
    End Select
        DrawEdge picLeft.hdc, LeftRect, BDR_SUNKENINNER, BF_TOPLEFT
        DrawEdge picLeft.hdc, LeftRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
        DrawEdge picRight.hdc, RightRect, BDR_SUNKENINNER, BF_TOPLEFT
        DrawEdge picRight.hdc, RightRect, BDR_SUNKENOUTER, BF_BOTTOMRIGHT
End Sub
Public Sub About()
Attribute About.VB_Description = "© Copyright 2001 Vincent Foster"
Attribute About.VB_UserMemId = -552
    frmAbout.Show 1, Me
End Sub
Public Property Get HTMLColor() As String
Dim HexRed As String
Dim HexGreen As String
Dim HexBlue As String

        HexRed = Hex(m_cRedValue)
    If Len(HexRed) = 1 Then HexRed = "0" & HexRed
    
        HexGreen = Hex(m_cGreenValue)
    If Len(HexGreen) = 1 Then HexGreen = "0" & HexGreen
    
        HexBlue = Hex(m_cBlueValue)
    If Len(HexBlue) = 1 Then HexBlue = "0" & HexBlue

    HTMLColor = "#" & HexRed & HexGreen & HexBlue
End Property
Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get cRampStyle() As RampStyleEnum
    cRampStyle = m_cRampStyle
End Property
Public Property Let cRampStyle(ByVal New_cRampStyle As RampStyleEnum)
    m_cRampStyle = New_cRampStyle
    PropertyChanged "cRampStyle"
    mnuRamp_Click m_cRampStyle
End Property
Private Function MakeGray(lngRed As Long, lngGreen As Long, lngBlue As Long) As Integer
    Dim Gray As Long
    If lngRed = 255 And lngGreen = 255 And lngBlue = 255 Then
    MakeGray = 255
    Exit Function
    End If
    
    lngRed = 0.3 * lngRed
    lngGreen = 0.59 * lngGreen
    lngBlue = 0.11 * lngBlue
    MakeGray = lngRed + lngGreen + lngBlue
    Select Case MakeGray
    Case 54
    MakeGray = 51
    Case 103
    MakeGray = 102
    Case 205
    MakeGray = 204
    Case Else
    End Select
End Function
Private Function MakeWebSafe(intRGB As Integer) As Integer
Select Case intRGB
Case 0 To 25
MakeWebSafe = 0
Case 26 To 76
MakeWebSafe = 51
Case 77 To 127
MakeWebSafe = 102
Case 128 To 178
MakeWebSafe = 153
Case 179 To 229
MakeWebSafe = 204
Case 230 To 255
MakeWebSafe = 255
End Select
End Function

