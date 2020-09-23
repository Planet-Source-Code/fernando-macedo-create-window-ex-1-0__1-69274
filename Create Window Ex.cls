VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hwnd As Long
Public Property Get Get_Window_ExStyle() As Long
Get_Window_ExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
End Property
Public Property Get Get_Window_Style() As Long
Get_Window_Style = GetWindowLong(hwnd, GWL_STYLE)
End Property
Public Property Get Get_Window_Height() As Integer
Dim lpRect As RECT
GetWindowRect hwnd, lpRect
Get_Window_Height = lpRect.Bottom - lpRect.Top
End Property
Public Property Get Get_Window_Left() As Integer
Dim lpRect As RECT
Dim m_Rect As RECT
GetWindowRect GetParent(hwnd), m_Rect
GetWindowRect hwnd, lpRect
Get_Window_Left = lpRect.Left - m_Rect.Left - 3
End Property
Public Property Get Get_Window_Top() As Integer
Dim lpRect As RECT
Dim m_Rect As RECT
GetWindowRect GetParent(hwnd), m_Rect
GetWindowRect hwnd, lpRect
Get_Window_Top = lpRect.Top - m_Rect.Top - 3
End Property
Public Property Get Get_Window_Child_ID() As Long
Get_Window_Child_ID = GetDlgCtrlID(hwnd)
End Property
Public Property Get Get_Window_Text() As String
Dim m_String As String * 255
Dim m_Len As Long
m_String = String(255, vbNullChar)
m_Len = GetWindowText(hwnd, m_String, 255)
Get_Window_Text = Left(m_String, m_Len)
End Property
Public Property Get Get_Window_ClassName() As String
Dim m_String As String * 255
Dim m_Len As Long
m_Len = GetClassName(hwnd, m_String, 255)
Get_Window_ClassName = Left(m_String, m_Len)
End Property
Public Property Get Get_Window_Width() As Integer
Dim lpRect As RECT
GetWindowRect hwnd, lpRect
Get_Window_Width = lpRect.Right - lpRect.Left
End Property
