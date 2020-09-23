VERSION 5.00
Object = "*\AJW_CoolForm.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4260
   ClientLeft      =   4410
   ClientTop       =   3630
   ClientWidth     =   6555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6555
   Begin JW_CoolFormVbp.JW_CoolForm JW_CoolForm1 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   556
      Caption         =   "CoolForm - ADMAX 2000"
      ForeColor       =   -2147483640
      FontSize        =   6,75
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   6435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' This code was written by Jim Williams (jim@admax.se)
' in order to enhance to VB-Form´s boring look.
' I needed a cool form for fixed forms, that hosted
' color-setups, forms that never should be rezised.
'
' Use it if you like it, and if you do - all I ask from you is
' that you send me an email (jim@admax.se) and tell me that
' you will use my code.
'
' (I have wrote other cool controls also, search for "jw_" inside Planet Source Code...)
'
' Enjoy this free code...


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub JW_CoolForm1_ClickMinimize()

End Sub


