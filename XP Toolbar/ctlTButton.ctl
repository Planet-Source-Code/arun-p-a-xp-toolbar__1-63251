VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlTButton 
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   4110
   ScaleWidth      =   4800
   Begin MSComctlLib.ImageList imgLst 
      Left            =   2400
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   3
      ImageHeight     =   38
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0149
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":01FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":034D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0497
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0551
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0699
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":071F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0860
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":090C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlTButton.ctx":0EE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer hLCheck 
      Interval        =   99
      Left            =   960
      Top             =   3240
   End
   Begin VB.PictureBox PicPlat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Image imgDRD 
         Height          =   45
         Left            =   3360
         Picture         =   "ctlTButton.ctx":0F6F
         Top             =   90
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Image imgRun 
         Height          =   390
         Left            =   45
         Stretch         =   -1  'True
         Top             =   90
         Width           =   390
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1635
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Image imgOvr 
      Height          =   495
      Left            =   240
      Top             =   3240
      Width           =   495
   End
   Begin VB.Image imgDis 
      Height          =   375
      Left            =   240
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image imgPic 
      Height          =   495
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
End
Attribute VB_Name = "ctlTButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim s_w As Long, s_h As Long
'Enum tlbButtonType
'    Normal = 0
'    Seperator = 1
'    DropDown = 2
'    PushButton = 3
'End Enum



Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type


'Default Property Values:
Const m_def_Theme = 0
'Const m_def_Theme = 0
'Const m_def_Theme = 0
Const m_def_Style = 0
Const m_def_Enabled = 1
'Property Variables:
Dim m_Theme As Integer
'Dim m_Theme As Integer
'Dim m_Theme As Integer
'Dim m_Spacing As Long
Dim m_Style As Integer
Dim m_Enabled As Boolean
'Event Declarations:
Event Click()
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOver(X As Long, Y As Long)
Event MouseOut(X As Long, Y As Long)

Dim down As Boolean

Public inMode As Integer

Private Sub DrawAll()
    Select Case m_Theme
        Case 0
            PicPlat.BackColor = RGB(243, 243, 238)
        Case 1
            PicPlat.BackColor = RGB(236, 239, 244)
    End Select
    If m_Style = 1 Then Exit Sub
        
    lblCaption.Visible = CBool(Len(lblCaption.Caption))
    
    Height = 570
    Width = CalcWidth
    
    If m_Style = 2 Then
        imgDRD.Left = Width - 8 * 15 - 10
        imgDRD.Top = (525 / 2)
    End If
    
    If lblCaption.Visible Then
        lblCaption.Left = IIf(imgPic.Picture <> 0, imgPic.Width, (Width - lblCaption.Width) / 2)
        lblCaption.Top = (Height - lblCaption.Height) / 2
    End If
    
    If m_Enabled Then
        If imgPic.Picture <> 0 Then imgRun.Picture = imgPic.Picture
    Else
        imgRun.Picture = IIf(imgDis.Picture <> 0, imgDis.Picture, imgPic.Picture)
    End If
    
    If lblCaption.Visible Then lblCaption.Refresh
    
    If inMode Then
        Highlight inMode, True
    End If
End Sub

Private Sub hLCheck_Timer()
    If m_Style = 1 Then hLCheck = False

    Dim lPt As POINTAPI, xhW As Long
    GetCursorPos& lPt
    xhW = WindowFromPoint(lPt.X, lPt.Y)
    If inMode = 0 And xhW = PicPlat.hWnd Then

        If down Then
            PicPlat.PaintPicture imgLst.ListImages(4 + m_Theme * 7).Picture, 0, 0
            PicPlat.PaintPicture imgLst.ListImages(5 + m_Theme * 7).Picture, 45, 0, Width - 90
            PicPlat.PaintPicture imgLst.ListImages(6 + m_Theme * 7).Picture, Width - 45, 0
        Else
            PicPlat.PaintPicture imgLst.ListImages(1 + m_Theme * 7).Picture, 0, 0
            PicPlat.PaintPicture imgLst.ListImages(2 + m_Theme * 7).Picture, 45, 0, Width - 90
            PicPlat.PaintPicture imgLst.ListImages(3 + m_Theme * 7).Picture, Width - 45, 0
        End If
        

        If imgOvr.Picture <> 0 Then
            imgRun.Picture = imgOvr.Picture
        ElseIf imgPic.Picture <> 0 Then
            imgRun.Picture = imgPic.Picture
        End If
        
        lblCaption.Refresh
        RaiseEvent MouseOver(lPt.X, lPt.Y)
        inMode = 1
    End If
    If inMode <> 0 And xhW <> PicPlat.hWnd And (down = False) Then
        PicPlat.Cls
        If imgPic.Picture <> 0 Then
            imgRun.Picture = imgPic.Picture
        End If
        RaiseEvent MouseOut(lPt.X, lPt.Y)
        inMode = 0
    End If
End Sub

Private Function CalcWidth() As Long
    If m_Style = 1 Then
        CalcWidth = 6 * 15
        Exit Function
    End If
    
    If lblCaption.Caption = "" Then
        CalcWidth = 570
    Else
        CalcWidth = lblCaption.Width + IIf(imgPic.Picture <> 0, imgPic.Width, 0) + 15 * 8
    End If
    
    If m_Style = 2 Then CalcWidth = CalcWidth + 91
    
    If m_Style = 2 And lblCaption = "" And imgPic.Picture = 0 Then
        CalcWidth = 210
    End If
End Function

Private Sub imgDRD_Click()
    If m_Enabled Then RaiseEvent Click
End Sub
Private Sub imgDRD_dblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub
'event handleres (passing)
Private Sub lblCaption_Click()
    If m_Enabled Then RaiseEvent Click
End Sub
Private Sub lblCaption_DblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub

Private Sub imgRun_Click()
    If m_Enabled Then RaiseEvent Click
End Sub
Private Sub imgRun_dblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub

Private Sub PicPlat_Click()
    If m_Enabled Then RaiseEvent Click
End Sub
Private Sub PicPlat_DblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub


Private Sub imgRun_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picplat_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgRun_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picplat_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picplat_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picplat_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub PicPlat_KeyDown(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub PicPlat_KeyPress(KeyAscii As Integer)
    If m_Enabled Then RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub PicPlat_KeyUp(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub
''''''''''''''''''''

Private Sub picplat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled = False Or m_Style = 1 Then Exit Sub
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = vbRightButton Then Exit Sub
    
    If m_Style = 3 And down Then
        picplat_MouseUp -1, Shift, X, Y
        Exit Sub
    End If
    
    
    
    imgRun.Top = imgRun.Top + 15
    imgRun.Left = imgRun.Left + 15
    
    'If lblCaption.Visible Then
        lblCaption.Top = lblCaption.Top + 15
        lblCaption.Left = lblCaption.Left + 15
        lblCaption.ForeColor = IIf(m_Style <> 3, vbWhite, vbBlack)
    'End If
    
    If m_Style = 3 Then
        PicPlat.PaintPicture imgLst.ListImages(15).Picture, 0, 0
        PicPlat.PaintPicture imgLst.ListImages(16).Picture, 45, 0, Width - 90
        PicPlat.PaintPicture imgLst.ListImages(17).Picture, Width - 45, 0
    Else
        PicPlat.PaintPicture imgLst.ListImages(4 + m_Theme * 7).Picture, 0, 0
        PicPlat.PaintPicture imgLst.ListImages(5 + m_Theme * 7).Picture, 45, 0, Width - 90
        PicPlat.PaintPicture imgLst.ListImages(6 + m_Theme * 7).Picture, Width - 45, 0
    End If
    If imgPic.Picture <> 0 Then
        'picplat.PaintPicture imgPic.Picture, IIf(lblCaption.Visible, 10, 90), 90, 26 * 15, 26 * 15
        imgRun.Picture = imgPic.Picture
    End If
    
    down = True
    inMode = 1
    If m_Style = 3 Then inMode = 2
End Sub

Private Sub picplat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled = False Or m_Style = 1 Then Exit Sub
    If Button <> -1 Then RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = vbRightButton Then Exit Sub
    If m_Style = 3 And Button <> -1 Then Exit Sub
    
    imgRun.Top = imgRun.Top - 15
    imgRun.Left = imgRun.Left - 15
    'If lblCaption.Visible Then
        lblCaption.Top = lblCaption.Top - 15
        lblCaption.Left = lblCaption.Left - 15
        lblCaption.ForeColor = vbBlack
    'End If

    
    PicPlat.PaintPicture imgLst.ListImages(1 + m_Theme * 7).Picture, 0, 0
    PicPlat.PaintPicture imgLst.ListImages(2 + m_Theme * 7).Picture, 45, 0, Width - 90
    PicPlat.PaintPicture imgLst.ListImages(3 + m_Theme * 7).Picture, Width - 45, 0
'    PicPlat.Cls
    
    If imgPic.Picture <> 0 Then
        'picplat.PaintPicture imgPic.Picture, IIf(lblCaption.Visible, 10, 90), 90, 26 * 15, 26 * 15
        imgRun.Picture = imgPic.Picture
    End If
    down = False
    inMode = 0
    If m_Style = 3 Then inMode = 1
End Sub

Private Sub UserControl_Paint()
    Height = 570
    Width = CalcWidth
    DrawAll
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Highlight 0, False
    
    m_Enabled = New_Enabled
    hLCheck.Enabled = New_Enabled
    If m_Style <> 1 And imgDis.Picture <> 0 Then
        'picplat.PaintPicture IIf(m_Enabled, imgPic.Picture, imgDis.Picture), IIf(lblCaption.Visible, 10, 90), 90, 26 * 15, 26 * 15
        imgRun.Picture = IIf(m_Enabled, imgPic.Picture, imgDis.Picture)
    End If
    If m_Style = 2 Then
        imgDRD.Visible = New_Enabled
    End If
    lblCaption.ForeColor = IIf(New_Enabled, vbBlack, RGB(157, 157, 161))
    
    PropertyChanged "Enabled"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     DrawAll
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgPic,imgPic,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = imgPic.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set imgPic.Picture = New_Picture
    imgRun.Visible = CBool(imgPic.Picture <> 0)
    Width = CalcWidth
    DrawAll
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption

    lblCaption.Visible = CBool(Len(New_Caption))
    
    imgRun.Left = IIf(lblCaption.Visible, 60, 90)
    DrawAll

    
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Style = m_def_Style
        
    down = False
'    m_Theme = m_def_Theme
'    m_Theme = m_def_Theme
    m_Theme = m_def_Theme
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "")
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    Set Picture = PropBag.ReadProperty("PictureOver", Nothing)
    Set Picture = PropBag.ReadProperty("PictureDisabled", Nothing)
    PicPlat.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
'    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
'    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("PictureOver", Picture, Nothing)
    Call PropBag.WriteProperty("PictureDisabled", Picture, Nothing)
    Call PropBag.WriteProperty("ToolTipText", PicPlat.ToolTipText, "")
'    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
'    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Style() As Integer
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As Integer)
    m_Style = New_Style
    
    Set PicPlat.Picture = Nothing
    imgDRD.Visible = False
    
    If m_Style = 1 Then
        imgRun.Visible = False
        Width = 6 * 15
        lblCaption.Visible = False
        PicPlat.Picture = imgLst.ListImages(7 * (1 + m_Theme)).Picture
        hLCheck = False
    ElseIf m_Style = 2 Then
        imgDRD.Visible = True
        Width = CalcWidth
    End If
    
    PropertyChanged "Style"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgOvr,imgOvr,-1,Picture
Public Property Get PictureOver() As Picture
Attribute PictureOver.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set PictureOver = imgOvr.Picture
End Property

Public Property Set PictureOver(ByVal New_PictureOver As Picture)
    Set imgOvr.Picture = New_PictureOver
    PropertyChanged "PictureOver"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgDis,imgDis,-1,Picture
Public Property Get PictureDisabled() As Picture
Attribute PictureDisabled.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set PictureDisabled = imgDis.Picture
End Property

Public Property Set PictureDisabled(ByVal New_PictureDisabled As Picture)
    Set imgDis.Picture = New_PictureDisabled
    PropertyChanged "PictureDisabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picplat,picplat,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = PicPlat.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    PicPlat.ToolTipText() = New_ToolTipText
    lblCaption.ToolTipText = New_ToolTipText
    imgRun.ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Highlight(hL As Integer, hLlock As Boolean) As Boolean
    If m_Enabled And (m_Style <> 1) Then
        imgRun.Top = 90
        imgRun.Left = IIf(lblCaption.Visible, 60, 90)
        
        'If lblCaption.Visible Then
            lblCaption.Left = IIf(imgPic.Picture <> 0, imgPic.Width, (Width - lblCaption.Width) / 2)
            lblCaption.Top = (Height - lblCaption.Height) / 2
            lblCaption.ForeColor = vbBlack
        'End If
        
        If hL = 1 Then
            PicPlat.PaintPicture imgLst.ListImages(1 + m_Theme * 7).Picture, 0, 0
            PicPlat.PaintPicture imgLst.ListImages(2 + m_Theme * 7).Picture, 45, 0, Width - 90
            PicPlat.PaintPicture imgLst.ListImages(3 + m_Theme * 7).Picture, Width - 45, 0
            
            If imgOvr.Picture <> 0 Then
                imgRun.Picture = imgOvr.Picture
            ElseIf imgPic.Picture <> 0 Then
                imgRun.Picture = imgPic.Picture
            End If
            inMode = 1
            down = CBool(m_Style <> 3)
            
        ElseIf hL = 2 Then
            imgRun.Top = imgRun.Top + 15
            imgRun.Left = imgRun.Left + 15
            If lblCaption.Visible Then
                lblCaption.Top = lblCaption.Top + 15
                lblCaption.Left = lblCaption.Left + 15
            End If
            lblCaption.ForeColor = IIf(m_Style <> 3, vbWhite, vbBlack)
            
        If m_Style = 3 Then
            PicPlat.PaintPicture imgLst.ListImages(15).Picture, 0, 0
            PicPlat.PaintPicture imgLst.ListImages(16).Picture, 45, 0, Width - 90
            PicPlat.PaintPicture imgLst.ListImages(17).Picture, Width - 45, 0
        Else
            PicPlat.PaintPicture imgLst.ListImages(4 + m_Theme * 7).Picture, 0, 0
            PicPlat.PaintPicture imgLst.ListImages(5 + m_Theme * 7).Picture, 45, 0, Width - 90
            PicPlat.PaintPicture imgLst.ListImages(6 + m_Theme * 7).Picture, Width - 45, 0
        End If
        If imgPic.Picture <> 0 Then imgRun.Picture = imgPic.Picture
            inMode = 2
            down = True
        Else
            PicPlat.Cls
            If imgPic.Picture <> 0 Then imgRun.Picture = imgPic.Picture
            inMode = 0
            down = False
        End If
    End If
    Highlight = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Theme() As Integer
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As Integer)
    m_Theme = New_Theme
    PropertyChanged "Theme"
End Property

