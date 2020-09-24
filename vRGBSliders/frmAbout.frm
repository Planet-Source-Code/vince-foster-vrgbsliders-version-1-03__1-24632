VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About vRGBSliders"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "vRGBSliders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   180
      Picture         =   "frmAbout.frx":038A
      Top             =   780
      Width           =   240
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: alienheretic@home.com"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   660
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   840
      Width           =   2235
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright © 2000 Vincent Foster"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Dim highlighted As Boolean
Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub Form_Load()
Me.Caption = "About " & App.Title
Label1.Caption = App.Title & "  Beta Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < lblEmail.Left Or Y < lblEmail.Top Or _
       X > lblEmail.Left + lblEmail.Width Or _
       Y > lblEmail.Top + lblEmail.Height _
    Then
        highlighted = False
        lblEmail.ForeColor = vbBlack
    Else
     highlighted = True
    lblEmail.ForeColor = vbBlue
    End If
End Sub
Private Sub lblEmail_Click()
sendemail "alienheretic@home.com"
End Sub
Public Sub sendemail(Email As String)
Dim Success As Long
Success = ShellExecute(0&, vbNullString, "mailto:" & Email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub
Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Form_MouseMove Button, Shift, X + lblEmail.Left, Y + lblEmail.Top
End Sub


