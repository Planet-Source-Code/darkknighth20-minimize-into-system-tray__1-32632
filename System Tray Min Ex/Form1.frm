VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Tray Example - by DarkKnight"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Minimize into the System Tray"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FF0000&
      Caption         =   "Mail Bomber"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF0000&
      Caption         =   "PW Cracker"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   75
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Example by Darkknight"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This shows u how to minimize a form into the system tray. I believe this method was in aim filter too. "
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
Me.show
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = " Example by DarkKnight " & vbNullChar
End With
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
Select Case msg
Case WM_LBUTTONUP
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
Me.show
Shell_NotifyIcon NIM_DELETE, nid
Case WM_LBUTTONDBLCLK
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
Me.show
Shell_NotifyIcon NIM_DELETE, nid
Case WM_RBUTTONUP
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
Me.show
Shell_NotifyIcon NIM_DELETE, nid
End Select
End Sub
Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Me.hide
Shell_NotifyIcon NIM_ADD, nid
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub
Private Sub Label7_Click()
Me.WindowState = vbMinimized
End Sub

