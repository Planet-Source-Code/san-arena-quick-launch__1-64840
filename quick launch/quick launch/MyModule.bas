Attribute VB_Name = "MyModule"
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public Function isTransparent(ByVal hWnd As Long) As Boolean
On Error Resume Next
Dim msg As Long
msg = GetWindowLong(hWnd, GWL_EXSTYLE)
If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
  isTransparent = True
Else
  isTransparent = False
End If
If Err Then
  isTransparent = False
End If
End Function

Public Function MakeTransparent(ByVal hWnd As Long, Perc As Integer) As Long
Dim msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  msg = GetWindowLong(hWnd, GWL_EXSTYLE)
  msg = msg Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, msg
  SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hWnd As Long) As Long
Dim msg As Long
On Error Resume Next
msg = GetWindowLong(hWnd, GWL_EXSTYLE)
msg = msg And Not WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, msg
SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then
  MakeOpaque = 2
End If
End Function

Public Sub CreateKey(Folder As String, Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value

End Sub

Public Sub CreateIntegerKey(Folder As String, Value As Integer)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value, "REG_DWORD"


End Sub

Public Function ReadKey(Value As String) As String

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
r = b.RegRead(Value)
ReadKey = r
End Function


Public Sub DeleteKey(Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("Wscript.Shell")
b.RegDelete Value
End Sub
Public Sub GrabForm(Frm As Form)
    ReleaseCapture
    SendMessage Frm.hWnd, &HA1, 2, ByVal 0&
CreateKey "HKCU\Software\Taskar\qlt", qlaunch.Top
CreateKey "HKCU\Software\Taskar\qll", qlaunch.Left
End Sub
