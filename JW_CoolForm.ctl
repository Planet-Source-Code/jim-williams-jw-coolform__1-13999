VERSION 5.00
Begin VB.UserControl JW_CoolForm 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LockControls    =   -1  'True
   ScaleHeight     =   300
   ScaleWidth      =   6210
   ToolboxBitmap   =   "JW_CoolForm.ctx":0000
   Begin VB.Image picImage 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   60
      Picture         =   "JW_CoolForm.ctx":0312
      Stretch         =   -1  'True
      ToolTipText     =   "For future use..."
      Top             =   35
      Width           =   240
   End
   Begin VB.Image picHide 
      Height          =   240
      Left            =   5640
      Picture         =   "JW_CoolForm.ctx":200C
      ToolTipText     =   "Minimize..."
      Top             =   35
      Width           =   240
   End
   Begin VB.Image picMinimize 
      Height          =   240
      Left            =   5400
      Picture         =   "JW_CoolForm.ctx":2596
      ToolTipText     =   "Help about this window..."
      Top             =   35
      Width           =   240
   End
   Begin VB.Image picClose 
      Height          =   240
      Left            =   5880
      Picture         =   "JW_CoolForm.ctx":2B20
      ToolTipText     =   "Close..."
      Top             =   35
      Width           =   240
   End
   Begin VB.Label CapText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CoolForm - ADMAX 2000"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Klicka och håll ner muspilen här, för att flytta fönstret..."
      Top             =   60
      Width           =   6210
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   780
      Left            =   0
      Picture         =   "JW_CoolForm.ctx":30AA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11385
   End
End
Attribute VB_Name = "JW_CoolForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private var1 As Long
Private var2 As Long
Private wStat As Long

Enum Aligns
    AlignLeft = 0
    AlignRight = 1
    AlignCenter = 2
End Enum

' API Calls
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' API Constants
Private Const GWL_STYLE = (-16)
Private Const LB_ADDSTRING = &H180
Private Const LB_FINDSTRING = &H18F
Private Const LB_RESETCONTENT = &H184
Private Const WS_DLGFRAME = &H400000    '3D Ram
Private Const WS_BORDER = &H800000      'Tunn svart linje
'Default Property Values:

'Const m_def_ForeColor = 0
'Const m_def_BackColor = &HFFFFFF
'Const m_def_ForeColor = vbWhite
Const m_def_Enabled = 0
'Const m_def_BackStyle = 0
'Const m_def_BorderStyle = 0
'Property Variables:

'Dim m_ForeColor As OLE_COLOR
'Dim m_BackColor As OLE_COLOR
'Dim m_ForeColor As OLE_COLOR
Dim m_Enabled As Boolean
'Dim m_Font As Font
'Dim m_BackStyle As Integer
'Dim m_BorderStyle As Integer
'Event Declarations:
Event DblClick() 'MappingInfo=CapText,CapText,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event ClickClose() 'MappingInfo=picClose,picClose,-1,Click
Attribute ClickClose.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event ClickHelp() 'MappingInfo=picMinimize,picMinimize,-1,Click
Event ClickHide() 'MappingInfo=picHide,picHide,-1,Click
Attribute ClickHide.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event ClickIcon() 'MappingInfo=pic2,pic2,-1,Click


Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=CapText,CapText,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=CapText,CapText,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=CapText,CapText,-1,MouseDown
'Event Click()
'Event DblClick()
'Event KeyDown(KeyCode As Integer, Shift As Integer)
'Event KeyPress(KeyAscii As Integer)
'Event KeyUp(KeyCode As Integer, Shift As Integer)
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'
'
'
'Private Sub pic2_Click()
'    RaiseEvent ClickIcon
'
'End Sub

Public Sub FormDrag()

    ReleaseCapture
    Call SendMessage(Parent.hWnd, &HA1, 2, 0&)
    
End Sub
Public Function FormsOnTop(OnTop As Boolean)

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Dim lState As Long
Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer

With Parent
    iLeft = .left / Screen.TwipsPerPixelX
    iTop = .top / Screen.TwipsPerPixelY
    iWidth = .Width / Screen.TwipsPerPixelX
    iHeight = .Height / Screen.TwipsPerPixelY
End With

If OnTop Then
    lState = HWND_TOPMOST
Else
    lState = HWND_NOTOPMOST
End If

Call SetWindowPos(Parent.hWnd, lState, iLeft, iTop, iWidth, iHeight, 0)

End Function






Private Sub picClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

picClose.left = picClose.left + 15
picClose.top = picClose.top + 15
DoEvents

End Sub

Private Sub picClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picClose.left = (Width - picClose.Width) - 50
picClose.top = 35
DoEvents
End Sub


Private Sub picHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picHide.left = picHide.left + 15
picHide.top = picHide.top + 15
DoEvents
End Sub

Private Sub picHide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picHide.left = picClose.left - 240
picHide.top = 35
DoEvents
End Sub


Private Sub picImage_Click()
    RaiseEvent ClickIcon
End Sub

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

picImage.left = picImage.left + 15
picImage.top = picImage.top + 15
DoEvents

End Sub


Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picImage.left = 60
picImage.top = 35
DoEvents
End Sub

Private Sub picMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMinimize.left = picMinimize.left + 15
picMinimize.top = picMinimize.top + 15
DoEvents
End Sub

Private Sub picMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picMinimize.left = picHide.left - 240
picMinimize.top = 35
DoEvents
End Sub


Private Sub UserControl_AmbientChanged(PropertyName As String)

If PropertyName = "BackColor" Then
    Parent.BackColor = &HE0E0E0
End If

End Sub

Private Sub UserControl_Initialize()

'3D-border
'SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_DLGFRAME

'Plain border
SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) Or WS_BORDER

End Sub
Public Function FormCenter()

Dim X As Long, Y As Long

X = Screen.Width / 2 - Parent.Width / 2
Y = Screen.Height / 2 - Parent.Height / 2

Parent.Move X, Y
End Function

Private Sub UserControl_Resize()
Image1.Width = Width
Image1.Height = Height

CapText.Width = Width
CapText.Height = Height

picClose.left = (Width - picClose.Width) - 50
picHide.left = picClose.left - 240
picMinimize.left = picHide.left - 240


End Sub
'
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=10,0,0,&hFFFFFF&
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = m_BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    m_BackColor = New_BackColor
'    PropertyChanged "BackColor"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=10,0,0,vbwhite
''Public Property Get ForeColor() As OLE_COLOR
''    ForeColor = m_ForeColor
''    CapText.ForeColor = ForeColor
''End Property
''
''Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
''    m_ForeColor = New_ForeColor
''    CapText.ForeColor = m_ForeColor
''    PropertyChanged "ForeColor"
''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=6,0,0,0
'Public Property Get Font() As Font
'    Set Font = m_Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set m_Font = New_Font
'    PropertyChanged "Font"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=7,0,0,0
''Public Property Get BackStyle() As Integer
''    BackStyle = m_BackStyle
''End Property
''
''Public Property Let BackStyle(ByVal New_BackStyle As Integer)
''    m_BackStyle = New_BackStyle
''    PropertyChanged "BackStyle"
''End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MemberInfo=7,0,0,0
''Public Property Get BorderStyle() As Integer
''    BorderStyle = m_BorderStyle
''End Property
''
''Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
''    m_BorderStyle = New_BorderStyle
''    PropertyChanged "BorderStyle"
''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=CapText,CapText,-1,ForeColor
'Public Property Get TextForeColor() As OLE_COLOR
'    TextForeColor = CapText.ForeColor
'End Property
'
'Public Property Let TextForeColor(ByVal New_TextForeColor As OLE_COLOR)
'    CapText.ForeColor() = New_TextForeColor
'    PropertyChanged "TextForeColor"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    Parent.ControlBox = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    CapText.Caption = PropBag.ReadProperty("Caption", "Enter your text here...")
    CapText.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    CapText.FontItalic = PropBag.ReadProperty("FontItalic", False)
    CapText.FontSize = PropBag.ReadProperty("FontSize", 8)
    CapText.FontBold = PropBag.ReadProperty("FontBold", False)
    Set CapText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    CapText.Alignment = PropBag.ReadProperty("AlignCaption", 2)
    Image1.BorderStyle = PropBag.ReadProperty("ImageBack", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000D)
    Set Picture = PropBag.ReadProperty("Icon", Nothing)

End Sub

Private Sub UserControl_Show()

var1 = Parent.Height
var2 = Parent.Width

Parent.BackColor = &HE0E0E0
Parent.BorderStyle = 1
Parent.ClipControls = False
Parent.Caption = ""
FormCenter
wStat = 1

'--- The following line should be NOT be marked during compilation!!! ---
'Call MouseCapture

End Sub

Private Sub UserControl_Terminate()
MouseRelease
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", CapText.Caption, "Enter your text here...")
    Call PropBag.WriteProperty("ForeColor", CapText.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("FontItalic", CapText.FontItalic, False)
    Call PropBag.WriteProperty("FontSize", CapText.FontSize, 8)
    Call PropBag.WriteProperty("FontBold", CapText.FontBold, False)
    Call PropBag.WriteProperty("Font", CapText.Font, Ambient.Font)
    Call PropBag.WriteProperty("AlignCaption", CapText.Alignment, 2)
    Call PropBag.WriteProperty("ImageBack", Image1.BorderStyle, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000D)
    Call PropBag.WriteProperty("Icon", Picture, Nothing)
    
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=CapText,CapText,-1,Caption
'Public Property Get TextCaption() As String
'    TextCaption = CapText.Caption
'End Property
'
'Public Property Let TextCaption(ByVal New_TextCaption As String)
'    CapText.Caption() = New_TextCaption
'    PropertyChanged "TextCaption"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapText,CapText,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = CapText.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    CapText.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=10,0,0,0
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = m_ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    m_ForeColor = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapText,CapText,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = CapText.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    CapText.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=CapText,CapText,-1,FontName
'Public Property Get FontName() As String
'    FontName = CapText.FontName
'End Property
'
'Public Property Let FontName(ByVal New_FontName As String)
'    CapText.FontName() = New_FontName
'    PropertyChanged "FontName"
'End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=CapText,CapText,-1,FontBold
''Public Property Get FontBold() As Boolean
''    FontBold = CapText.FontBold
''End Property
''
''Public Property Let FontBold(ByVal New_FontBold As Boolean)
''    CapText.FontBold() = New_FontBold
''    PropertyChanged "FontBold"
''End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapText,CapText,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = CapText.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    CapText.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapText,CapText,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = CapText.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    CapText.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapText,CapText,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = CapText.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    CapText.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapText,CapText,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = CapText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set CapText.Font = New_Font
    PropertyChanged "Font"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=CapText,CapText,-1,BorderStyle
'Public Property Get BorderStyle() As Boolean
'    BorderStyle = CapText.BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Boolean)
'
'    If New_BorderStyle = True Then
'        CapText.BorderStyle() = 1
'    Else
'        CapText.BorderStyle() = 0
'    End If
'
'    PropertyChanged "BorderStyle"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=CapText,CapText,-1,Alignment
'Public Property Get TextAlign() As Integer
'    TextAlign = CapText.Alignment
'End Property
'
'Public Property Let TextAlign(ByVal New_TextAlign As Integer)
'    CapText.Alignment() = New_TextAlign
'    PropertyChanged "TextAlign"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=CapText,CapText,-1,Alignment
Public Property Get AlignCaption() As Aligns
Attribute AlignCaption.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    AlignCaption = CapText.Alignment
End Property

Public Property Let AlignCaption(ByVal New_AlignCaption As Aligns)
    CapText.Alignment() = New_AlignCaption
    PropertyChanged "AlignCaption"
End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=Image1,Image1,-1,Enabled
''Public Property Get ImageBack() As Boolean
''    ImageBack = Image1.Enabled
''End Property
''
''Public Property Let ImageBack(ByVal New_ImageBack As Boolean)
''    Image1.Enabled() = New_ImageBack
''    PropertyChanged "ImageBack"
''End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=Image1,Image1,-1,BorderStyle
'Public Property Get ImageBack() As Integer
'    ImageBack = Image1.BorderStyle
'End Property
'
'Public Property Let ImageBack(ByVal New_ImageBack As Integer)
'    Image1.BorderStyle() = New_ImageBack
'    PropertyChanged "ImageBack"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Private Sub CapText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CapText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub CapText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    FormDrag
    If wStat = 1 Then Call MouseCapture
End Sub

Private Sub picClose_Click()
    RaiseEvent ClickClose
    Unload Parent
End Sub

Private Sub picMinimize_Click()
RaiseEvent ClickHelp


End Sub

Private Sub picHide_Click()
RaiseEvent ClickHide
    
If Parent.Height = Height + 30 Then
    Parent.Height = var1
    Parent.Width = var2
    FormsOnTop False
    Call MouseCapture
    wStat = 1
Else
    Parent.Height = Height + 30
    FormsOnTop True
    Call MouseRelease
    wStat = 0
End If

End Sub

Private Sub CapText_DblClick()
    RaiseEvent DblClick
    picHide_Click
End Sub
Public Function MouseCapture()

Dim client As RECT
Dim upperleft As POINT

GetClientRect Parent.hWnd, client
upperleft.X = client.left
upperleft.Y = client.top
ClientToScreen Parent.hWnd, upperleft
OffsetRect client, upperleft.X, upperleft.Y
ClipCursor client

End Function
Public Function MouseRelease()

ClipCursor ByVal 0&

End Function




