VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlXPToolBar 
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   ScaleHeight     =   3930
   ScaleWidth      =   5400
   ToolboxBitmap   =   "ctlXPToolBar.ctx":0000
   Begin MSComctlLib.ImageList imgGB 
      Left            =   3960
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   41
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlXPToolBar.ctx":0314
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlXPToolBar.ctx":0378
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrAutoM 
      Interval        =   99
      Left            =   240
      Top             =   3240
   End
   Begin VB.PictureBox picBase 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin prjXPToolBar.ctlTButton ToolButtons 
         Height          =   570
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "ctlXPToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_Theme = 0
'Const m_def_Theme = 0
Const m_def_Count = -1
Const m_def_Spacing = 12
Const m_def_AutoManage = 0
'Const m_def_Count = -1
Const m_def_Enabled = 1
'Property Variables:
Dim m_Theme As Integer
'Dim m_Theme As Integer
Dim m_Count As Integer
Dim m_Spacing As Long
Dim m_AutoManage As Boolean
'Dim m_Count As Integer
Dim m_Enabled As Boolean
Dim m_Font As Font
'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

Event ChildDblClick(sKey As String)
Event ChildKeyDown(sKey As String, KeyCode As Integer, Shift As Integer)
Event ChildKeyPress(sKey As String, KeyAscii As Integer)
Event ChildKeyUp(sKey As String, KeyCode As Integer, Shift As Integer)
Event ChildMouseDown(sKey As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ChildMouseMove(sKey As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ChildMouseUp(sKey As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ChildClick(sKey As String, X As Long, Y As Long)
Event ChildOver(sKey As String, X As Long, Y As Long)
Event ChildOut(sKey As String, X As Long, Y As Long)

Dim mgd As Boolean
Dim enStr As String
''''''''
Private Sub picBase_Click()
    If m_Enabled Then RaiseEvent Click
End Sub
Private Sub picBase_dblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub

Private Sub picBase_KeyDown(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub



Private Sub picBase_KeyPress(KeyAscii As Integer)
    If m_Enabled Then RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub picBase_KeyUp(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picBase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'''''''''''''''''''
Private Sub tmrAutoM_Timer()
    tmrAutoM = m_AutoManage
    If Not mgd Then ManageNow
    mgd = True
End Sub


''''''''''''child events ''''''''''''''''
Private Sub ToolButtons_Click(Index As Integer)
    If m_Enabled Then RaiseEvent ChildClick(ToolButtons(Index).Tag, ToolButtons(Index).Left, ToolButtons(Index).Top + ToolButtons(Index).Height)
End Sub

Private Sub ToolButtons_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If m_Enabled Then RaiseEvent ChildKeyDown(ToolButtons(Index).Tag, KeyCode, Shift)
End Sub

Private Sub ToolButtons_KeyPress(Index As Integer, KeyAscii As Integer)
    If m_Enabled Then RaiseEvent ChildKeyPress(ToolButtons(Index).Tag, KeyAscii)
End Sub

Private Sub ToolButtons_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If m_Enabled Then RaiseEvent ChildKeyUp(ToolButtons(Index).Tag, KeyCode, Shift)
End Sub

Private Sub ToolButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then RaiseEvent ChildMouseDown(ToolButtons(Index).Tag, Button, Shift, X, Y)
End Sub

Private Sub ToolButtons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then RaiseEvent ChildMouseMove(ToolButtons(Index).Tag, Button, Shift, X, Y)
End Sub

Private Sub ToolButtons_MouseOut(Index As Integer, X As Long, Y As Long)
    If m_Enabled Then RaiseEvent ChildOut(ToolButtons(Index).Tag, X, Y)
End Sub

Private Sub ToolButtons_MouseOver(Index As Integer, X As Long, Y As Long)
    If m_Enabled Then RaiseEvent ChildOver(ToolButtons(Index).Tag, X, Y)
End Sub

Private Sub ToolButtons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then RaiseEvent ChildMouseUp(ToolButtons(Index).Tag, Button, Shift, X, Y)
End Sub
'''''''''''
Private Sub UserControl_Paint()
    picBase.BackColor = RGB(243, 243, 238)
    picBase.Height = 645
    picBase.PaintPicture imgGB.ListImages(m_Theme + 1).Picture, 0, 0, Width
    Height = 645
    
    ToolButtons(0).Top = (Height - ToolButtons(0).Height) / 2 - 10
End Sub

Private Sub UserControl_Resize()
    picBase.Width = Width
    picBase.PaintPicture imgGB.ListImages(m_Theme + 1).Picture, 0, 0, Width
    Height = 645
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "ppgXPToolBar"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    If m_Enabled = False Then
        SaveEnb True
    Else
        RestoreEnb
    End If
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     For i = 0 To m_Count
        ToolButtons(i).Refresh
     Next
     ManageNow
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
'    m_Count = m_def_Count
    m_AutoManage = m_def_AutoManage
    m_Spacing = m_def_Spacing
    m_Count = m_def_Count
'    m_Theme = m_def_Theme
    m_Theme = m_def_Theme
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_Count = PropBag.ReadProperty("Count", m_def_Count)
    m_AutoManage = PropBag.ReadProperty("AutoManage", m_def_AutoManage)
    m_Spacing = PropBag.ReadProperty("Spacing", m_def_Spacing)
    m_Count = PropBag.ReadProperty("Count", m_def_Count)
'    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
'    Call PropBag.WriteProperty("Count", m_Count, m_def_Count)
    Call PropBag.WriteProperty("AutoManage", m_AutoManage, m_def_AutoManage)
    Call PropBag.WriteProperty("Spacing", m_Spacing, m_def_Spacing)
    Call PropBag.WriteProperty("Count", m_Count, m_def_Count)
'    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9
Public Function Buttons(strkey As String) As Object
    If m_Count < 0 Then Err.Raise 101, , "No Buttons Stored"

    i = GetIndex(strkey)
    If i < 0 Then
        Err.Raise 102, "Buttons()<-", "Invalid Button Key"
        Exit Function
    End If
    
    Set Buttons = ToolButtons(i)
    mgd = False
End Function

Private Function GetIndex(strkey As String) As Integer
    For i = 0 To m_Count
        If ToolButtons(i).Tag = strkey Then
            GetIndex = i
            Exit Function
        End If
    Next
    If i > m_Count Then GetIndex = -2
End Function

Private Function SaveEnb(Optional mb As Boolean = False)
    enStr = ""
    For i = 0 To m_Count
        enStr = enStr & IIf(ToolButtons(i).Enabled, "1", "0") & CStr(ToolButtons(i).inMode)
        If mb Then ToolButtons(i).Enabled = False
    Next
End Function
Private Function RestoreEnb()
    For i = 0 To m_Count
        ToolButtons(i).Enabled = CBool(Mid(enStr, i * 2 + 1, 1))
        ToolButtons(i).Highlight CInt(Mid(enStr, i * 2 + 2, 1)), True
    Next
End Function
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get Count() As Integer
'    Count = m_Count
'End Property
'
'Public Property Let Count(ByVal New_Count As Integer)
'    m_Count = New_Count
'    PropertyChanged "Count"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Remove(strkey As String) As Boolean
    i = GetIndex(strkey)
    If i < 0 Then
        Err.Raise 102, "Remove()->", "Invalid Button Key"
        Exit Function
    End If
    RestoreEnb
    For j = i To m_Count - 1
        With ToolButtons(j)
           .Style = ToolButtons(j + 1).Style
           .Caption = ToolButtons(j + 1).Caption
           .Tag = ToolButtons(j + 1).Tag
           .Enabled = ToolButtons(j + 1).Enabled
           .ToolTipText = ToolButtons(j + 1).ToolTipText
            Set .Picture = ToolButtons(j + 1).Picture
            Set .PictureDisabled = ToolButtons(j + 1).PictureDisabled
            Set .PictureOver = ToolButtons(j + 1).PictureOver
           .Left = ToolButtons(j + 1).Left
           .Top = ToolButtons(j + 1).Top
           .Width = ToolButtons(j + 1).Width
           .Height = ToolButtons(j + 1).Height
           .Refresh
        End With
    Next j
    
    Unload ToolButtons(m_Count)
    m_Count = m_Count - 1
    If m_AutoManage Then ManageNow
    SaveEnb
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function ManageNow() As Boolean
    For i = 0 To m_Count
        If i = 0 Then
            ToolButtons(0).Left = 15
            ToolButtons(0).Top = (Height - ToolButtons(0).Height) / 2 - 10
        Else
            ToolButtons(i).Left = ToolButtons(i - 1).Left + ToolButtons(i - 1).Width + m_Spacing
            ToolButtons(i).Top = ToolButtons(0).Top
        End If
    Next i
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Add(strkey As String) As Boolean
    If GetIndex(strkey) <> -2 Then
        Err.Raise 103, "Add", "Invalid Keyname (multiples)"
        Exit Function
    End If
    
    m_Count = m_Count + 1
    If m_Count = 0 Then
        ToolButtons(0).Left = 15
        ToolButtons(0).Top = (Height - ToolButtons(0).Height) / 2 - 10
    Else
        Load ToolButtons(m_Count)
        ToolButtons(m_Count).Left = ToolButtons(m_Count - 1).Left + ToolButtons(m_Count - 1).Width + m_Spacing
        ToolButtons(m_Count).Top = ToolButtons(0).Top
    End If
    ToolButtons(m_Count).Theme = m_Theme
    ToolButtons(m_Count).Style = 0
    ToolButtons(m_Count).Visible = True
    ToolButtons(m_Count).Enabled = True
    ToolButtons(m_Count).Tag = strkey
    Set ToolButtons(m_Count).Picture = Nothing
    ToolButtons(m_Count).Caption = ""
    ToolButtons(m_Count).ToolTipText = ""
    SaveEnb
End Function





'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoManage() As Boolean
    AutoManage = m_AutoManage
End Property

Public Property Let AutoManage(ByVal New_AutoManage As Boolean)
    m_AutoManage = New_AutoManage
    tmrAutoM = m_AutoManage
    PropertyChanged "AutoManage"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function Highlight(sKey As String, Optional hLit As Integer = 1, Optional LockAfter As Boolean = False) As Boolean
    i = GetIndex(sKey)
    If i < 0 Then
        Err.Raise 102, "Highlight->", "Invalid Button Key"
        Exit Function
    End If
    Highlight = ToolButtons(i).Highlight(hLit, LockAfter)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,12
Public Property Get Spacing() As Long
    Spacing = m_Spacing
End Property

Public Property Let Spacing(ByVal New_Spacing As Long)
    m_Spacing = New_Spacing
    If m_AutoManage Then ManageNow
    PropertyChanged "Spacing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,2,0
Public Property Get Count() As Integer
Attribute Count.VB_MemberFlags = "400"
    Count = m_Count
End Property

Public Property Let Count(ByVal New_Count As Integer)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Count = New_Count
    PropertyChanged "Count"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get Theme() As Integer
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As Integer)
    If Ambient.UserMode Then Err.Raise 382
    m_Theme = New_Theme
    PropertyChanged "Theme"
End Property

