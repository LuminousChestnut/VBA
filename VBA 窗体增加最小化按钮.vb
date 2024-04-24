VBA 窗体增加最小化按钮！！
' via: https://cloud.tencent.com/developer/article/1468329, 一线编程
'全局声明

Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long '获取窗口样式

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long '查找当前窗口句柄

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16) '设置窗口样式

Private Const WS_MINIMIZEBOX As Long = &H20000 '最小化

'==========================================================

'窗体UserForm的初始化

Private Sub UserForm_Initialize()

Dim hWndForm As Long

Dim IStyle As Long

hWndForm = FindWindow("ThunderDFrame", Me.Caption)  '获取窗口句柄

IStyle = GetWindowLong(hWndForm, GWL_STYLE) '获取当前标题栏样式

IStyle = IStyle Or WS_MINIMIZEBOX '设置最小化按钮

SetWindowLong hWndForm, GWL_STYLE, IStyle  '显示最小化按钮

End Sub
