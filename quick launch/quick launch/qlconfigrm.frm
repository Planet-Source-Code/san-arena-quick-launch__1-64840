VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form config 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   " ‰ŸÌ„« "
   ClientHeight    =   4875
   ClientLeft      =   3735
   ClientTop       =   2730
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "qlconfigrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "qlconfigrm.frx":164A
   RightToLeft     =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Close"
      Height          =   375
      Left            =   1440
      MaskColor       =   &H8000000F&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Apply"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "qlconfigrm.frx":22DB
      Left            =   3840
      List            =   "qlconfigrm.frx":22EB
      TabIndex        =   20
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse..."
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txa3 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Browse..."
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txa4 
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txt4 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox txt3 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse..."
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txa2 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   -120
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.*"
      DialogTitle     =   "›«Ì· Ì« ÅÊ‘Â „Ê—œ ‰Ÿ— ŒÊœ —« «‰ Œ«» ‰„«ÌÌœ"
      Filter          =   "*.*"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse..."
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txa1 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Run in Startup"
      Height          =   255
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Skin:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image img1 
      Height          =   255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Four Button Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Four Button Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Three Button Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Three Button Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Two Button Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Two Button Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Button Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Button Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim bytRegion(287) As Byte
Dim nBytes As Long
Private Sub Command1_Click()
On Error Resume Next
cd1.ShowOpen
txa1.Text = cd1.FileName
End Sub

Private Sub Command2_Click()
On Error Resume Next

cd1.ShowOpen
txa2.Text = cd1.FileName
End Sub

Private Sub Command3_Click()
On Error Resume Next

cd1.ShowOpen
txa3.Text = cd1.FileName
End Sub

Private Sub Command4_Click()
On Error Resume Next
cd1.ShowOpen
txa4.Text = cd1.FileName
End Sub

Private Sub Command5_Click()
If List1.ListIndex = -1 Then List1.ListIndex = ReadKey("HKCU\Software\Taskar\qlstyle") - 1
CreateKey ("HKCU\Software\Taskar\qlstyle"), List1.ListIndex + 1
qlaunch.loadimg List1.ListIndex + 1
CreateKey ("HKCU\Software\Taskar\qlbtntxt1"), txt1
CreateKey ("HKCU\Software\Taskar\qlbtntxt2"), txt2
CreateKey ("HKCU\Software\Taskar\qlbtntxt3"), txt3
CreateKey ("HKCU\Software\Taskar\qlbtntxt4"), txt4
CreateKey ("HKCU\Software\Taskar\qlbtntxa1"), txa1
CreateKey ("HKCU\Software\Taskar\qlbtntxa2"), txa2
CreateKey ("HKCU\Software\Taskar\qlbtntxa3"), txa3
CreateKey ("HKCU\Software\Taskar\qlbtntxa4"), txa4
qlaunch.Loadtext
If Check1.Value = 1 Then
CreateKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Taskarql", App.Path & "\QuickLaunch.exe"
Else
DeleteKey "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Taskarql"
End If
Me.Hide
End Sub

Private Sub Command6_Click()
Me.Hide
End Sub

Private Sub Form_Load()

Dim rgnMain As Long

nBytes = 288

LoadBytes

rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
SetWindowRgn Me.hWnd, rgnMain, True

If Not ReadKey("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Taskarql") = "" Then
Check1.Value = 1
Else
Check1.Value = 0
End If
img1.Picture = LoadPicture(App.Path & "\quicklaunch\style" & ReadKey("HKCU\Software\Taskar\qlstyle") & ".gif")
txt1 = ReadKey("HKCU\Software\Taskar\qlbtntxt1")
txt2 = ReadKey("HKCU\Software\Taskar\qlbtntxt2")
txt3 = ReadKey("HKCU\Software\Taskar\qlbtntxt3")
txt4 = ReadKey("HKCU\Software\Taskar\qlbtntxt4")
txa1 = ReadKey("HKCU\Software\Taskar\qlbtntxa1")
txa2 = ReadKey("HKCU\Software\Taskar\qlbtntxa2")
txa3 = ReadKey("HKCU\Software\Taskar\qlbtntxa3")
txa4 = ReadKey("HKCU\Software\Taskar\qlbtntxa4")
End Sub
Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 16
bytRegion(13) = 1
bytRegion(24) = 70
bytRegion(25) = 1
bytRegion(28) = 70
bytRegion(29) = 1
bytRegion(32) = 6
bytRegion(40) = 63
bytRegion(41) = 1
bytRegion(44) = 1
bytRegion(48) = 5
bytRegion(52) = 1
bytRegion(56) = 64
bytRegion(57) = 1
bytRegion(60) = 2
bytRegion(64) = 3
bytRegion(68) = 2
bytRegion(72) = 66
bytRegion(73) = 1
bytRegion(76) = 3
bytRegion(80) = 2
bytRegion(84) = 3
bytRegion(88) = 67
bytRegion(89) = 1
bytRegion(92) = 4
bytRegion(96) = 1
bytRegion(100) = 4
bytRegion(104) = 67
bytRegion(105) = 1
bytRegion(108) = 5
bytRegion(112) = 1
bytRegion(116) = 5
bytRegion(120) = 68
bytRegion(121) = 1
bytRegion(124) = 6
bytRegion(132) = 6
bytRegion(136) = 69
bytRegion(137) = 1
bytRegion(140) = 8
bytRegion(148) = 8
bytRegion(152) = 70
bytRegion(153) = 1
bytRegion(156) = 61
bytRegion(157) = 1
bytRegion(164) = 61
bytRegion(165) = 1
bytRegion(168) = 69
bytRegion(169) = 1
bytRegion(172) = 63
bytRegion(173) = 1
bytRegion(176) = 1
bytRegion(180) = 63
bytRegion(181) = 1
bytRegion(184) = 68
bytRegion(185) = 1
bytRegion(188) = 64
bytRegion(189) = 1
bytRegion(192) = 2
bytRegion(196) = 64
bytRegion(197) = 1
bytRegion(200) = 67
bytRegion(201) = 1
bytRegion(204) = 65
bytRegion(205) = 1
bytRegion(208) = 2
bytRegion(212) = 65
bytRegion(213) = 1
bytRegion(216) = 66
bytRegion(217) = 1
bytRegion(220) = 66
bytRegion(221) = 1
bytRegion(224) = 3
bytRegion(228) = 66
bytRegion(229) = 1
bytRegion(232) = 65
bytRegion(233) = 1
bytRegion(236) = 67
bytRegion(237) = 1
bytRegion(240) = 5
bytRegion(244) = 67
bytRegion(245) = 1
bytRegion(248) = 64
bytRegion(249) = 1
bytRegion(252) = 68
bytRegion(253) = 1
bytRegion(256) = 6
bytRegion(260) = 68
bytRegion(261) = 1
bytRegion(264) = 63
bytRegion(265) = 1
bytRegion(268) = 69
bytRegion(269) = 1
bytRegion(272) = 9
bytRegion(276) = 69
bytRegion(277) = 1
bytRegion(280) = 60
bytRegion(281) = 1
bytRegion(284) = 70
bytRegion(285) = 1
End Sub
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 0 Then
Check1.Value = 1
Else
Check1.Value = 0
End If
End Sub

Private Sub List1_Click()
img1.Picture = LoadPicture(App.Path & "\quicklaunch\style" & List1.ListIndex + 1 & ".gif")
End Sub
