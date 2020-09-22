Attribute VB_Name = "mPopUpMenuMod"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public MenuChanged As Boolean
Public MenuReturn() As String
Public MenuForms() As mPopUpMenu
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long



Public Function PopMenu(MenuData() As clsMenuItem, Optional Index = 0, Optional X = -1, Optional Y = -1, Optional BackColor As OLE_COLOR = vbButtonFace) As Form
If Index = 0 And Not MenuForms(0) Is Nothing Then
MenuChanged = True
End If
DoEvents
If UBound(MenuReturn) < Index Or UBound(MenuForms) < Index Then
DimIt:
ReDim Preserve MenuReturn(0 To Index)
ReDim Preserve MenuForms(0 To Index)
End If
If UBound(MenuReturn) < Index Or UBound(MenuForms) < Index Then GoTo DimIt
MenuReturn(Index) = -1
If Not MenuForms(Index) Is Nothing Then Unload MenuForms(Index)
Set MenuForms(Index) = New mPopUpMenu
If X < 0 Or Y < 0 Then
Dim c As POINTAPI
GetCursorPos c
If X < 0 Then X = c.X * Screen.TwipsPerPixelX
If Y < 0 Then Y = c.Y * Screen.TwipsPerPixelY
End If
Load MenuForms(Index)
MenuForms(Index).SetMenuItems MenuData()
MenuForms(Index).bFace.BackColor = BackColor
RePosMenu MenuForms(Index), X, Y
MenuForms(Index).mIndex = Index
MenuForms(Index).Show
MenuForms(Index).ZOrder 0
Set PopMenu = MenuForms(Index)
End Function

Public Function RGBRed(RGBCol As Long) As Integer
    If RGBCol > 0 Then RGBRed = RGBCol And &HFF Else RGBRed = 0
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
    If RGBCol > 0 Then RGBGreen = ((RGBCol And &H100FF00) / &H100) Else RGBGreen = 0
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
    If RGBCol > 0 Then RGBBlue = (RGBCol And &HFF0000) / &H10000 Else RGBBlue = 0
End Function

Public Sub RePosMenu(tFrm As Form, X, Y)
tFrm.RefreshMenu
If tFrm.Width + X + Screen.TwipsPerPixelX > Screen.Width Then
tFrm.Left = X - tFrm.Width - Screen.TwipsPerPixelX
Else
tFrm.Left = X + Screen.TwipsPerPixelX
End If
If tFrm.Height + Y + Screen.TwipsPerPixelY > Screen.Height Then
tFrm.Top = Y - tFrm.Height - Screen.TwipsPerPixelY
Else
tFrm.Top = Y + Screen.TwipsPerPixelY
End If
End Sub
