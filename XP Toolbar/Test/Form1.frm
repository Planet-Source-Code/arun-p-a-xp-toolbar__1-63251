VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\prjXPToolBar.vbp"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin prjXPToolBar.ctlXPToolBar ctlXPToolBar1 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoManage      =   -1  'True
      Theme           =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0256
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":04CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSomen 
      Caption         =   "mnuSomen"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "blaa"
      End
      Begin VB.Menu mnu2 
         Caption         =   "blaa2"
      End
      Begin VB.Menu mnu3 
         Caption         =   "booo"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnmods As Boolean, blnexes As Boolean



Private Sub Command2_Click()
    ctlXPToolBar1.Enabled = Not ctlXPToolBar1.Enabled
End Sub

Private Sub ctlXPToolBar1_ChildMouseDown(sKey As String, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case sKey
        Case "kS_Mods"
            blnmods = Not blnmods
        Case "kS_Exes"
            blnexes = Not blnexes
            If blnexes = False Then blnmods = False
            ctlXPToolBar1.Buttons("kS_Mods").Enabled = blnexes
    End Select
End Sub

Private Sub ctlXPToolBar1_ChildOut(sKey As String, X As Long, Y As Long)
    If sKey = "kSearch" Then
        ctlXPToolBar1.Highlight "kP", False, False
    ElseIf sKey = "kP" Then
        ctlXPToolBar1.Highlight "kSearch", False, False
    End If
End Sub

Private Sub ctlXPToolBar1_ChildOver(sKey As String, X As Long, Y As Long)
    If sKey = "kSearch" Then
        ctlXPToolBar1.Highlight "kP", True, True
    ElseIf sKey = "kP" Then
        ctlXPToolBar1.Highlight "kSearch", True, True
    End If
End Sub


Private Sub Form_Load()
    With ctlXPToolBar1
        .Add "kS_Exes"
        .Buttons("kS_Exes").Style = 3
        .Buttons("kS_Exes").Picture = Form1.ImageList1.ListImages(1).Picture
        .Buttons("kS_Exes").Caption = "Processes"
        .Highlight "kS_Exes", 2
        
        .Add "kS_Mods"
        .Buttons("kS_Mods").Style = 3
        .Buttons("kS_Mods").Picture = Form1.ImageList1.ListImages(2).Picture
        .Buttons("kS_Mods").Picturedisabled = Form1.ImageList1.ListImages(3).Picture
        .Buttons("kS_Mods").Caption = "Modules"
        .Highlight "kS_Mods", 2
        
        .Add "kS_Mod2"
        .Buttons("kS_Mod2").Style = 0
        .Buttons("kS_Mod2").Picture = Form1.ImageList1.ListImages(2).Picture
        .Buttons("kS_Mod2").Picturedisabled = Form1.ImageList1.ListImages(3).Picture
        .Buttons("kS_Mod2").Caption = "Modules2"
        .Highlight "kS_Mod2", 1, True
        
        
    End With
    blnmods = True
    blnexes = True
End Sub
