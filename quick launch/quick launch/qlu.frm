VERSION 5.00
Begin VB.Form qlu 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1005
   ClientLeft      =   7875
   ClientTop       =   2535
   ClientWidth     =   1470
   Icon            =   "qlu.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   1470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "qlu.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "qlu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
config.Show
End Sub

Private Sub Command2_Click()
qlhelp.Show
End Sub

Private Sub Form_Resize()
Command1.Top = Me.Height - 850
End Sub
