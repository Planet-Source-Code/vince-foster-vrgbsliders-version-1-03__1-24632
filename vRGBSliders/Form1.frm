VERSION 5.00
Object = "{2D7FA805-6BC7-11D5-A6FD-0020780DD8F0}#3.0#0"; "vColorControls.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vRBGSliders"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin vColorControls.vRGBSliders vRGBSliders1 
      Height          =   1545
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2725
      BorderStyle     =   4
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Show
vRGBSliders1.About
End Sub
Private Sub vRGBSliders1_Change()
'Me.Caption = "RGB(" & vRGBSliders1.cRedValue & "," & vRGBSliders1.cGreenValue & "," & vRGBSliders1.cBlueValue & ")"
'Me.Caption = "vRBGSliders - " & vRGBSliders1.CurrentColor
Me.Caption = vRGBSliders1.HTMLColor
End Sub
