VERSION 5.00
Begin VB.Form FrmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test of AnimatedCursorViewer ActiveX Control"
   ClientHeight    =   2085
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1005
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "&Load from resource"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "g"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "Stop"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton CmdAction 
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      ToolTipText     =   "Play"
      Top             =   840
      Width           =   375
   End
   Begin PrjTest.AnimatedCursorViewer AnimatedCursorViewer1 
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Teste"
      Top             =   120
      Width           =   735
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Menu mnuPopTest 
      Caption         =   "popTest"
      Visible         =   0   'False
      Begin VB.Menu mnuPopSub 
         Caption         =   "&Play"
         Index           =   0
      End
      Begin VB.Menu mnuPopSub 
         Caption         =   "&Stop"
         Index           =   1
      End
      Begin VB.Menu mnuPopSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopSub 
         Caption         =   "&Load from resource"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AnimatedCursorViewer1_Change()
 Label1.Caption = "State = " & IIf(AnimatedCursorViewer1.State = eACSPaused, "Paused", "Playing") & ", DrawFocus = " & AnimatedCursorViewer1.DrawFocus
 CmdAction(0).Enabled = Not (AnimatedCursorViewer1.State = eACSPlaying)
 CmdAction(1).Enabled = Not (AnimatedCursorViewer1.State = eACSPaused)
 mnuPopSub(0).Enabled = Not (AnimatedCursorViewer1.State = eACSPlaying)
 mnuPopSub(1).Enabled = Not (AnimatedCursorViewer1.State = eACSPaused)
End Sub

Private Sub AnimatedCursorViewer1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case vbKeyF2
   Call CmdAction_Click(0)
  Case vbKeyF5
   Call CmdAction_Click(1)
  Case vbKeyF9
   Call CmdAction_Click(2)
 End Select
End Sub

Private Sub AnimatedCursorViewer1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
   PopupMenu mnuPopTest
 End If
End Sub

Private Sub CmdAction_Click(Index As Integer)
 Select Case Index
  Case 0
   AnimatedCursorViewer1.Play
   CmdAction(0).Enabled = False
   CmdAction(1).Enabled = True
   
  Case 1
   AnimatedCursorViewer1.Pause
   CmdAction(0).Enabled = True
   CmdAction(1).Enabled = False
   
  Case 2
   Static LastID As Long
    Select Case LastID
     Case 101
      LastID = 102
      GoSub PlayCursor
     Case 102
      LastID = 103
      GoSub PlayCursor
     Case 103
      LastID = 101
      GoSub PlayCursor
     Case Else
      LastID = 101
      GoSub PlayCursor
    End Select
 End Select
Exit Sub


PlayCursor:
    If AnimatedCursorViewer1.LoadFromResource(LastID) = False Then
     MsgBox "Error on load resource file !", 16
     Exit Sub
    Else
     AnimatedCursorViewer1.Play
     Text1.Text = "ID:" & LastID & vbCrLf & "Using:" & AnimatedCursorViewer1.Filename
    End If
  Exit Sub
End Sub

Private Sub File1_Click()
 AnimatedCursorViewer1.Filename = File1.Path & IIf(Right(File1.Path, 1) <> "\", "\", "") & File1.Filename
 If AnimatedCursorViewer1.AutoPlay = False Then AnimatedCursorViewer1.Play
End Sub

Private Sub Form_Load()
 File1.Path = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "cursors"
 File1.Pattern = "*.ani"
 File1.Refresh
 Call AnimatedCursorViewer1_Change
End Sub

Private Sub mnuPopSub_Click(Index As Integer)
 Select Case Index
  Case 0
   Call CmdAction_Click(0)
   
  Case 1
   Call CmdAction_Click(1)
   
  Case 3
   Call CmdAction_Click(2)
 End Select
End Sub
