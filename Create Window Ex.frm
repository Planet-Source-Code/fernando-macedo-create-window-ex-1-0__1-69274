VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Window Ex"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   Icon            =   "Create Window Ex.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   675
      TabIndex        =   1
      Top             =   158
      Width           =   9750
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DragIcon        =   "Create Window Ex.frx":0E42
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mouse_Over As Boolean

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
Private Sub Form_Load()

    With Picture1
        .Left = 4
        .Top = 4
        .Height = 32
        .Width = 32
    End With

    Set m_Window = New Class1
    Picture1.Picture = Picture1.DragIcon
    Me.MouseIcon = Picture1.DragIcon
    
    Dim m_Top
    m_Top = SetWindowPos(Form1.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub
Private Sub Create_Window_Ex(hwnd As Long)

    If IsChild(GetParent(hwnd), hwnd) Then
    m_Window.hwnd = hwnd
    Dim lpRect As RECT
    Dim m_Mdc As Long
    
    GetWindowRect hwnd, lpRect
    m_Mdc = GetDC(0)
    Call DrawFocusRect(m_Mdc, lpRect)
    Call InflateRect(lpRect, -1, -1)
    Call DrawFocusRect(m_Mdc, lpRect)
    Call InflateRect(lpRect, -1, -1)
    Call DeleteDC(m_Mdc)

    Text1.Text = "m_hWnd = " _
    & "CreateWindowEx(" _
    & "&H" & Hex(m_Window.Get_Window_ExStyle) _
    & ", " & Chr(34) & m_Window.Get_Window_ClassName & Chr(34) _
    & ", " & Chr(34) & m_Window.Get_Window_Text & Chr(34) & ", " _
    & "&H" & Hex(m_Window.Get_Window_Style) & ", " _
    & m_Window.Get_Window_Left & ", " _
    & m_Window.Get_Window_Top - 19 & ", " _
    & m_Window.Get_Window_Width & ", " _
    & m_Window.Get_Window_Height & ", " _
    & "Me.hWnd" & ", " _
    & "0&" & ", " _
    & "0&" & ", " _
    & "&H" & Hex(m_Window.Get_Window_Child_ID) _
    & ")"
    
    Else
    Text1.Text = ""
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = vbCustom
    Set Picture1.Picture = Nothing
    Mouse_Over = True
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Mouse_Over = False Then Exit Sub

Static lpPoint As POINTAPI
GetCursorPos lpPoint
If lpPoint.x = x And lpPoint.y = y Then Exit Sub
x = lpPoint.x
y = lpPoint.y
On Local Error Resume Next

' HWND
Static hwnd As Long
Call GetCursorPos(lpPoint) ' Get cursor position
hwnd = WindowFromPoint(x, y) ' Get window cursor is over
Create_Window_Ex hwnd

End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture1.Picture = Picture1.DragIcon
    Me.MousePointer = vbDefault
    Mouse_Over = False
End Sub


