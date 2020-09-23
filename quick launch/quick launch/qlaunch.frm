VERSION 5.00
Begin VB.Form qlaunch 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Quick Launch"
   ClientHeight    =   450
   ClientLeft      =   3795
   ClientTop       =   2190
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "qlaunch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   100
      Width           =   1455
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   100
      Width           =   1455
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   100
      Width           =   1455
   End
   Begin VB.CommandButton utilitybtn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   100
      Width           =   615
   End
   Begin VB.CommandButton closebtn 
      BackColor       =   &H00FF0000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7130
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   100
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   75
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   100
      Width           =   1455
   End
End
Attribute VB_Name = "qlaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim bytRegion(143) As Byte
Dim nBytes As Long
Dim hiding
Dim updowning
Dim showing
Dim widths
Dim tops




Private Sub closebtn_Click()
End
End Sub

Private Sub cmd1_Click()
ShellExecute 0&, "Open", ReadKey("HKCU\Software\Taskar\qlbtntxa1"), "", vbNullString, 1
End Sub

Private Sub cmd1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub

Private Sub cmd2_Click()
ShellExecute 0&, "Open", ReadKey("HKCU\Software\Taskar\qlbtntxa2"), "", vbNullString, 1

End Sub

Private Sub cmd2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub

Private Sub cmd3_Click()
ShellExecute 0&, "Open", ReadKey("HKCU\Software\Taskar\qlbtntxa3"), "", vbNullString, 1

End Sub

Private Sub cmd3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub

Private Sub cmd4_Click()
ShellExecute 0&, "Open", ReadKey("HKCU\Software\Taskar\qlbtntxa4"), "", vbNullString, 1
End Sub

Private Sub cmd4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub



Private Sub Form_Load()
Dim rgnMain As Long
nBytes = 144
LoadBytes
rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
SetWindowRgn Me.hWnd, rgnMain, True
If Not ReadKey("HKCU\Software\Taskar\qlt") = "" Then
Me.Top = ReadKey("HKCU\Software\Taskar\qlt")
Me.Left = ReadKey("HKCU\Software\Taskar\qll")
End If
showing = 0
If ReadKey("HKCU\Software\Taskar\qlbtntxt1") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxt1"), ""
If ReadKey("HKCU\Software\Taskar\qlbtntxt2") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxt2"), ""
If ReadKey("HKCU\Software\Taskar\qlbtntxt3") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxt3"), ""
If ReadKey("HKCU\Software\Taskar\qlbtntxt4") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxt4"), ""
If ReadKey("HKCU\Software\Taskar\qlbtntxa1") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxa1"), App.Path & "\df.exe"
If ReadKey("HKCU\Software\Taskar\qlbtntxa2") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxa2"), App.Path & "\df.exe"
If ReadKey("HKCU\Software\Taskar\qlbtntxa3") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxa3"), App.Path & "\df.exe"
If ReadKey("HKCU\Software\Taskar\qlbtntxa4") = "" Then CreateKey ("HKCU\Software\Taskar\qlbtntxa4"), App.Path & "\df.exe"
Loadtext
If ReadKey("HKCU\Software\Taskar\qlstyle") = "" Then
CreateKey ("HKCU\Software\Taskar\qlstyle"), 1
loadimg 1
Else
loadimg ReadKey("HKCU\Software\Taskar\qlstyle")
End If
End Sub
Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 7
bytRegion(12) = 112
bytRegion(24) = 244
bytRegion(25) = 1
bytRegion(28) = 30
bytRegion(32) = 3
bytRegion(40) = 241
bytRegion(41) = 1
bytRegion(44) = 1
bytRegion(48) = 2
bytRegion(52) = 1
bytRegion(56) = 242
bytRegion(57) = 1
bytRegion(60) = 2
bytRegion(64) = 1
bytRegion(68) = 2
bytRegion(72) = 244
bytRegion(73) = 1
bytRegion(76) = 3
bytRegion(84) = 3
bytRegion(88) = 244
bytRegion(89) = 1
bytRegion(92) = 27
bytRegion(96) = 1
bytRegion(100) = 27
bytRegion(104) = 243
bytRegion(105) = 1
bytRegion(108) = 28
bytRegion(112) = 2
bytRegion(116) = 28
bytRegion(120) = 242
bytRegion(121) = 1
bytRegion(124) = 29
bytRegion(128) = 3
bytRegion(132) = 29
bytRegion(136) = 241
bytRegion(137) = 1
bytRegion(140) = 30
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
GrabForm qlaunch
End Sub



Private Sub Timer1_Timer()
If Not showing = 1 Then
qlu.Left = (qlaunch.Left) + 5820
If Not Val(hiding) > 220 Then
hiding = Val(hiding) + 5
End If
MakeTransparent qlu.hWnd, Val(hiding)
If Not qlu.Height > 1000 Then qlu.Height = qlu.Height + 20
If Not qlu.Top > 450 + qlaunch.Top Then qlu.Top = qlu.Top
If qlu.Top > 430 + qlaunch.Top And qlu.Height > 980 Then
showing = 1
Timer1.Enabled = False
End If
End If
End Sub

Private Sub Timer2_Timer()
qlu.Left = (qlaunch.Left) + 5820
config.Left = qlaunch.Left
config.Top = Me.Top + 550
If showing = 1 Then
qlu.Top = 450 + qlaunch.Top
Else
If Timer1.Enabled = False Then qlu.Top = qlaunch.Top
End If
End Sub


Private Sub utilitybtn_Click()
If showing = 0 Then
MakeTransparent qlu.hWnd, 0
Timer1.Enabled = True
qlu.Top = 450 + qlaunch.Top
qlu.Height = 1
qlu.Show
Else
Do While hiding > 5
hiding = hiding - 5
MakeTransparent qlu.hWnd, Val(hiding)
Loop
qlu.Hide
showing = 0
End If
End Sub

Private Sub utilitybtn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
End Sub

Function loadimg(stylenumber)
qlaunch.Picture = LoadPicture(App.Path & "\quicklaunch\bgoftask" & stylenumber & ".bmp")
cmd1.Picture = LoadPicture(App.Path & "\quicklaunch\upbtn" & stylenumber & ".bmp")
cmd2.Picture = LoadPicture(App.Path & "\quicklaunch\upbtn" & stylenumber & ".bmp")
cmd3.Picture = LoadPicture(App.Path & "\quicklaunch\upbtn" & stylenumber & ".bmp")
cmd4.Picture = LoadPicture(App.Path & "\quicklaunch\upbtn" & stylenumber & ".bmp")
utilitybtn.Picture = LoadPicture(App.Path & "\quicklaunch\upbtn" & stylenumber & ".bmp")
cmd1.DownPicture = LoadPicture(App.Path & "\quicklaunch\dwbtn" & stylenumber & ".bmp")
cmd2.DownPicture = LoadPicture(App.Path & "\quicklaunch\dwbtn" & stylenumber & ".bmp")
cmd3.DownPicture = LoadPicture(App.Path & "\quicklaunch\dwbtn" & stylenumber & ".bmp")
cmd4.DownPicture = LoadPicture(App.Path & "\quicklaunch\dwbtn" & stylenumber & ".bmp")
utilitybtn.DownPicture = LoadPicture(App.Path & "\quicklaunch\dwbtn" & stylenumber & ".bmp")
End Function
Function Loadtext()
cmd1.Caption = ReadKey("HKCU\Software\Taskar\qlbtntxt1")
cmd2.Caption = ReadKey("HKCU\Software\Taskar\qlbtntxt2")
cmd3.Caption = ReadKey("HKCU\Software\Taskar\qlbtntxt3")
cmd4.Caption = ReadKey("HKCU\Software\Taskar\qlbtntxt4")
End Function
