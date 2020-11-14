VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1班点名神器"
   ClientHeight    =   4620
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3120
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "置顶窗口"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Caption         =   "更少"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7800
      Top             =   3600
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "滚动"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "一键抽奖/人"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "重置"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   3
      Text            =   "3"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "抽奖"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim a(1 To 51) As String
Dim p(1 To 51) As Boolean
Dim rtn
Dim wg As Integer

Private Sub Check1_Click()
 If Check1 = 1 Then
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE '置顶
Else
'取消窗口在顶层
SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE Or SWP_NOMOVE '不置顶
End If
End Sub

Private Sub Command1_Click()
Randomize
Command1.Caption = "抽奖"
Timer1.Enabled = False
way1:
k = Int(Rnd() * 51 + 1)
If p(k) = False Then
Label1.Caption = a(k)
List1.AddItem a(k)
p(k) = True
Else
If p(1) = True And p(2) = True And p(3) = True And p(4) = True And p(5) = True And p(6) = True And p(7) = True And p(8) = True And p(9) = True And p(10) = True And p(11) = True And p(12) = True And p(13) = True And p(14) = True And p(15) = True And p(16) = True And p(17) = True And p(18) = True And p(19) = True And p(20) = True And p(21) = True And p(22) = True And p(23) = True And p(24) = True And p(25) = True And p(26) = True And p(27) = True And p(28) = True And p(29) = True And p(30) = True And p(31) = True And p(32) = True And p(33) = True And p(34) = True And p(35) = True And p(36) = True And p(37) = True And p(38) = True And p(39) = True And p(40) = True And p(41) = True And p(42) = True And p(43) = True And p(44) = True And p(45) = True And p(46) = True And p(47) = True And p(48) = True And p(49) = True And p(50) = True And p(51) = True Then
Label1.Caption = "请重置"
Else
GoTo way1
End If
End If
End Sub

Private Sub Command2_Click()
For i = 1 To 51
p(i) = flase
Next i
List1.Clear
End Sub

Private Sub Command3_Click()
Randomize
For i = 1 To Val(Text1.Text)
way2:
k = Int(Rnd() * 51 + 1)
If p(k) = False Then
List1.AddItem a(k)
p(k) = True
Else
If p(1) = True And p(2) = True And p(3) = True And p(4) = True And p(5) = True And p(6) = True And p(7) = True And p(8) = True And p(9) = True And p(10) = True And p(11) = True And p(12) = True And p(13) = True And p(14) = True And p(15) = True And p(16) = True And p(17) = True And p(18) = True And p(19) = True And p(20) = True And p(21) = True And p(22) = True And p(23) = True And p(24) = True And p(25) = True And p(26) = True And p(27) = True And p(28) = True And p(29) = True And p(30) = True And p(31) = True And p(32) = True And p(33) = True And p(34) = True And p(35) = True And p(36) = True And p(37) = True And p(38) = True And p(39) = True And p(40) = True And p(41) = True And p(42) = True And p(43) = True And p(44) = True And p(45) = True And p(46) = True And p(47) = True And p(48) = True And p(49) = True And p(50) = True And p(51) = True Then
Label1.Caption = "请重置"
Else
GoTo way2
End If
End If
Next i
End Sub

Private Sub Command4_Click()
Command1.Caption = "停止滚动并抽奖"
Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
If wg = 0 Then
Form1.Height = 2400
Command5.Caption = "更多"
Command4.Visible = False
Command3.Visible = False
Command2.Visible = False
Text1.Visible = False
List1.Visible = False
wg = 1
Exit Sub
End If
If wg = 1 Then
Form1.Height = 5205
Command5.Caption = "更少"
Command4.Visible = True
Command3.Visible = True
Command2.Visible = True
Text1.Visible = True
List1.Visible = True
wg = 0
Exit Sub
End If
End Sub

Private Sub Form_Load()
For i = 1 To 51
p(i) = flase
Next i
Check1 = 1
Command1.BackColor = RGB(255, 204, 108)
'*****
For k = 1 To 51
a(k) = "姓名" & Val(k)
Next k
'*****
End Sub

Private Sub Label1_DblClick()
Label1.Caption = "Au制作"
End Sub

Private Sub Timer1_Timer()
l = Int(Rnd() * 51 + 1)
Label1.Caption = a(l)
End Sub
