VERSION 5.00
Begin VB.UserControl AnimatedCursorViewer 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   PropertyPages   =   "AnimatedCursorViewer.ctx":0000
   ScaleHeight     =   585
   ScaleWidth      =   645
   ToolboxBitmap   =   "AnimatedCursorViewer.ctx":0026
End
Attribute VB_Name = "AnimatedCursorViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const SS_ICON As Long = &H3
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const STM_SETIMAGE As Long = &H172
Private Const IMAGE_CURSOR As Long = &H2

Public Enum eAnimatedCursorAppearanceConstants
 eACAFlat = 0
 eACA3D = 1
End Enum

Public Enum eAnimatedCursorBorderConstants
 eACBNone = 0
 eACBFixedSingle = 1
End Enum

Public Enum eAnimatedCursorStateConstants
 eACSPlaying
 eACSPaused
End Enum

Private Type ANICURSOR
 m_hCursor As Long
 m_hWnd As Long
End Type

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OSGetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal BufferLength As Long, ByVal Result As String) As Long
Private Declare Function OSGetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal FilePath As String, ByVal Prefix As String, ByVal wUnique As Long, ByVal TempFileName As String) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long


Private mAutoPlay As Boolean
Private mAniCursorObject As ANICURSOR
Private mFilename As String
Private mState As eAnimatedCursorStateConstants
Private mDrawFocus As Boolean

Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Resize()

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
 MsgBox "Animated Cursor Viewer UserControl." & vbCrLf & "Developed by Mauricio Cunha" & vbCrLf & "E-mail: mcunha98@terra.com.br" & vbCrLf & "Homepage: http://www.mcunha98.cjb.net", , "Animated Cursor Viewer"
End Sub

Public Property Get Appearance() As eAnimatedCursorAppearanceConstants
 Appearance = UserControl.Appearance
End Property
Public Property Let Appearance(ByVal NewValue As eAnimatedCursorAppearanceConstants)
 UserControl.Appearance = NewValue
 PropertyChanged "Appearance"
End Property

Public Property Get AutoPlay() As Boolean
 AutoPlay = mAutoPlay
End Property
Public Property Let AutoPlay(ByVal NewValue As Boolean)
 mAutoPlay = NewValue
 PropertyChanged "AutoPlay"
End Property

Public Property Get BackColor() As OLE_COLOR
 BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
 UserControl.BackColor = NewValue
 PropertyChanged "BackColor"
 RaiseEvent Change
End Property

Public Property Get BorderStyle() As eAnimatedCursorBorderConstants
 BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal NewValue As eAnimatedCursorBorderConstants)
 UserControl.BorderStyle = NewValue
 PropertyChanged "BorderStyle"
 RaiseEvent Change
End Property

Public Property Get DrawFocus() As Boolean
 DrawFocus = mDrawFocus
End Property
Public Property Let DrawFocus(ByVal NewValue As Boolean)
 mDrawFocus = NewValue
 PropertyChanged "DrawFocus"
End Property

Public Property Get Enabled() As Boolean
 Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
 UserControl.Enabled = NewValue
 PropertyChanged "Enabled"
End Property

Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
 Filename = mFilename
End Property
Public Property Let Filename(ByVal NewValue As String)
 mFilename = NewValue
 PropertyChanged "Filename"
 If AutoPlay = True And Dir(mFilename) <> "" Then Call Play
 RaiseEvent Change
End Property

Public Property Get MouseIcon() As StdPicture
 Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal NewValue As StdPicture)
 Set UserControl.MouseIcon = NewValue
 PropertyChanged "MouseIcon"
 RaiseEvent Change
End Property

Public Property Get MousePointer() As MousePointerConstants
 MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
 UserControl.MousePointer = NewValue
 PropertyChanged "MousePointer"
 RaiseEvent Change
End Property

Public Property Get State() As eAnimatedCursorStateConstants
 State = mState
End Property

Private Sub UserControl_EnterFocus()
 If DrawFocus = True Then
  Dim myRECT As RECT
  myRECT.Top = 1
  myRECT.Left = 1
  myRECT.Bottom = UserControl.ScaleHeight / Screen.TwipsPerPixelY - 1
  myRECT.Right = UserControl.ScaleWidth / Screen.TwipsPerPixelX - 1
  DrawFocusRect UserControl.hdc, myRECT
 End If
End Sub

Private Sub UserControl_ExitFocus()
 If DrawFocus = True Then UserControl.Cls
End Sub

Private Sub UserControl_Initialize()
 BorderStyle = eACBFixedSingle
 Filename = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Appearance = PropBag.ReadProperty("Appearance", 1)
 AutoPlay = PropBag.ReadProperty("AutoPlay", True)
 BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
 BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
 DrawFocus = PropBag.ReadProperty("DrawFocus", True)
 Enabled = PropBag.ReadProperty("Enabled", True)
 Filename = PropBag.ReadProperty("Filename", "")
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
 Call PropBag.WriteProperty("AutoPlay", mAutoPlay, True)
 Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
 Call PropBag.WriteProperty("DrawFocus", mDrawFocus, True)
 Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
 Call PropBag.WriteProperty("Filename", mFilename, "")
 Call PropBag.WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Public Function Play() As Boolean
  Pause
  
  If Trim(mFilename) = "" Then
   Play = False
   Exit Function
  End If
   
  If Dir(mFilename) = "" Then
   Play = False
   Err.Raise -92831, "AnimatedCursorViewer", "File not found !"
   Exit Function
  End If
  
  With mAniCursorObject
     .m_hCursor = LoadCursorFromFile(mFilename)
     If .m_hCursor Then
        .m_hWnd = CreateWindowEx(0, "Static", "", WS_CHILD Or WS_VISIBLE Or SS_ICON, ByVal 20, ByVal 20, 0, 0, UserControl.hwnd, 0, App.hInstance, ByVal 0)
        If .m_hWnd Then
           SendMessage .m_hWnd, STM_SETIMAGE, IMAGE_CURSOR, ByVal .m_hCursor
           SetWindowPos .m_hWnd, 0, 2, 2, 0, 0, SWP_NOZORDER Or SWP_NOSIZE
           mState = eACSPlaying
           Play = True
        Else
           DestroyCursor .m_hCursor
           mState = eACSPaused
           Play = False
        End If
     End If
  End With
  RaiseEvent Change
End Function

Public Sub Pause()
  With mAniCursorObject
     If .m_hCursor Then If DestroyCursor(.m_hCursor) Then .m_hCursor = 0
     If IsWindow(.m_hWnd) Then If DestroyWindow(.m_hWnd) Then .m_hWnd = 0
  End With
  mState = eACSPaused
  RaiseEvent Change
End Sub

Private Sub UserControl_Click()
 RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
 RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
 RaiseEvent Resize
End Sub

Private Function GetTempFile() As String
Dim sFilePath As String
Dim sTempResult As String
Dim lCharCount As Long
Const MAX_RETURN = 3000

  sTempResult = Space(MAX_RETURN)
  lCharCount = OSGetTempPath(MAX_RETURN, sTempResult)
  sFilePath = Left(sTempResult, lCharCount)
  sTempResult = Space(MAX_RETURN)
  lCharCount = OSGetTempFilename(sFilePath, "ani", 0, sTempResult)
  GetTempFile = Left(sTempResult, lCharCount)
End Function

Public Function LoadFromResource(ResourceID As Long) As Boolean
On Error GoTo err1

Dim FNumber As Integer
Dim DllBuffer() As Byte
Dim sTempFile As String

sTempFile = GetTempFile
DllBuffer = LoadResData(ResourceID, "CUSTOM")
FNumber = FreeFile

Open sTempFile For Binary Access Write As #FNumber
 Put #FNumber, , DllBuffer
Close #FNumber

Filename = sTempFile
LoadFromResource = True
Exit Function


err1:
 LoadFromResource = False
 Exit Function
End Function
