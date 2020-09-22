VERSION 5.00
Begin VB.Form mPopUpMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   3240
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   2880
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   2520
      Top             =   2280
   End
   Begin VB.PictureBox bFace 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   1920
      ScaleHeight     =   435
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox HiF 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   1800
      ScaleHeight     =   435
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox HB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      Height          =   495
      Left            =   1680
      ScaleHeight     =   435
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   1920
   End
End
Attribute VB_Name = "mPopUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mWindow As Long, MenuItems() As clsMenuItem, HoverItem As Integer, HoverY As Single
Public mIndex As Integer, LastItem As Integer, ItemClicked As Boolean

Public Sub SetMenuItems(Items() As clsMenuItem)
MenuItems = Items
RefreshMenu
'Me.SetFocus
'SetActiveWindow Me.hWnd
mWindow = Me.hwnd
'Timer1.Enabled = True
End Sub

Public Sub RefreshMenu()
Me.Cls
Dim Item As clsMenuItem, LastParent As clsMenuItem, Area As RECT, i, tHeight
Dim H1 As Long, H2 As Long, HF As Long, BC As Long, S1 As Long, S2 As Long, sF As Long, _
FC As Long, uDC As Long, bWidth, rWidth
bWidth = Me.Width
rWidth = 0
Area.Left = 0
Area.Right = Me.ScaleWidth
Area.Top = 0
uDC = Me.hdc
H1 = CLng(GetSetting(App.Title, "/", "HeaderGradientStart", 0))
H2 = CLng(GetSetting(App.Title, "/", "HeaderGradientEnd", &HFF0000))
HF = CLng(GetSetting(App.Title, "/", "HeaderGradientForeColor", &HDDCCCC))
S1 = CLng(GetSetting(App.Title, "/", "GradientStart", &H4B4239))
S2 = CLng(GetSetting(App.Title, "/", "GradientEnd", &HFF0000))
sF = CLng(GetSetting(App.Title, "/", "GradientForeColor", &HC0C0C0))
FC = CLng(GetSetting(App.Title, "/", "NormalColor", &H440000))
tHeight = Me.TextHeight("|") / 2
i = LBound(MenuItems)
Do
  Set Item = MenuItems(i)
  Select Case Item.Style
  Case 0 'header -22
    Set LastParent = Item
    If Item.Visible = True Then
      Item.LastTop = Area.Top
      Area.Bottom = Area.Top + 22
      DrawHGrad Area, H1, H2
      Me.Line (Area.Left, Area.Top)-(Area.Right, Area.Bottom), &H404040, B
      Me.Line (Area.Left + 1, Area.Top + 1)-(Area.Right - 1, Area.Bottom - 1), &HFFFFFF, B
      SetTextColor uDC, HF
      TextOut uDC, Area.Left + 6, Area.Top + (11 - tHeight), Item.Caption, Len(Item.Caption)
      Area.Top = Area.Bottom + 1
      If Me.TextWidth(Item.Caption) + 16 > rWidth Then rWidth = Me.TextWidth(Item.Caption) + 16
    End If
  Case 1 'sub header -18
    Set LastParent = Item
    If Item.Visible = True Then
      Item.LastTop = Area.Top
      Area.Bottom = Area.Top + 18
      If Item.Selected = False Then
      DrawHGrad Area, S1, S2
      SetTextColor uDC, sF
      Else
      FillRect Area.Top, Area.Bottom - 1, HB.Point(1, 1)
      SetTextColor uDC, HiF.Point(1, 1)
      End If
      TextOut uDC, Area.Left + 6, Area.Top + (9 - tHeight), Item.Caption, Len(Item.Caption)
      If Item.Opened = True Then
        DrawDownArrow uDC, Area.Right - 16, Area.Top + 7, H1
      Else
        DrawRightArrow uDC, Area.Right - 15, Area.Top + 5, H1
      End If
      Area.Top = Area.Bottom
      If Me.TextWidth(Item.Caption) + 28 > rWidth Then rWidth = Me.TextWidth(Item.Caption) + 28
    End If
  Case 2 'popout item -18
      If LastParent Is Nothing Then GoTo DoWithSkipage2
      If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage2 Else GoTo Arg2
DoWithSkipage2:
      If Item.Visible = True Then
      Item.LastTop = Area.Top
      Area.Bottom = Area.Top + 18
      If Item.Selected = False Then
      FillRect Area.Top, Area.Bottom, bFace.Point(1, 1)
      SetTextColor uDC, FC
      Else
      FillRect Area.Top, Area.Bottom, HB.Point(1, 1)
      SetTextColor uDC, HiF.Point(1, 1)
      End If
      TextOut uDC, Area.Left + 28, Area.Top + (9 - tHeight), Item.Caption, Len(Item.Caption)
      If Not Item.Icon Is Nothing Then Me.PaintPicture Item.Icon, Area.Left + 6, Area.Top + 1
      DrawRightArrow uDC, Area.Right - 11, Area.Top + 6, H1
      Area.Top = Area.Bottom + 1
      If Me.TextWidth(Item.Caption) + 45 > rWidth Then rWidth = Me.TextWidth(Item.Caption) + 45
    End If
Arg2:
  Case 3 'normal item -18
    If LastParent Is Nothing Then GoTo DoWithSkipage1
    If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage1 Else GoTo Arg1
DoWithSkipage1:
    If Item.Visible = True Then
      Item.LastTop = Area.Top
      Area.Bottom = Area.Top + 18
      If Item.Selected = False Then
      FillRect Area.Top, Area.Bottom, bFace.Point(1, 1)
      SetTextColor uDC, FC
      Else
      FillRect Area.Top, Area.Bottom, HB.Point(1, 1)
      SetTextColor uDC, HiF.Point(1, 1)
      End If
      TextOut uDC, Area.Left + 28, Area.Top + (9 - tHeight), Item.Caption, Len(Item.Caption)
      If Not Item.Icon Is Nothing Then Me.PaintPicture Item.Icon, Area.Left + 6, Area.Top + 1
      Area.Top = Area.Bottom + 1
      If Me.TextWidth(Item.Caption) + 34 > rWidth Then rWidth = Me.TextWidth(Item.Caption) + 34
    End If
Arg1:
  End Select
  i = i + 1
Loop Until i >= UBound(MenuItems)
If rWidth < 100 Then rWidth = 100
Me.Width = (rWidth + 6) * Screen.TwipsPerPixelX
Me.Height = (Area.Top + 6) * Screen.TwipsPerPixelY
If Me.Width <> bWidth Then RefreshMenu
Me.Refresh
End Sub

Private Sub FillRect(tStart, tEnd, Color)
Dim i
i = tStart
Do
DrawHLine i, Color
i = i + 1
Loop Until i > tEnd
End Sub

Private Sub DrawHLine(Y, Color)
Dim uDC
uDC = Me.hdc
SetPixel uDC, 0, Y, Color
If Me.ScaleWidth > 0 Then
StretchBlt uDC, 1, Y, Me.ScaleWidth - 1, 1, uDC, 0, Y, 1, 1, vbSrcCopy
End If
End Sub

Private Sub Form_GotFocus()
Me.ZOrder 0
Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mIndex > 0 Then
MenuForms(mIndex - 1).Timer2.Enabled = False
MenuForms(mIndex - 1).Timer3.Enabled = False
End If
Dim tI As Integer
tI = OverItem(Y)
i = LBound(MenuItems)
Do
MenuItems(i).Selected = i = tI
i = i + 1
Loop Until i >= UBound(MenuItems)
RefreshMenu
If Timer2.Enabled = True Then
If HoverItem <> tI And mIndex < UBound(MenuReturn) Then
Timer2.Enabled = False
Timer3.Enabled = True
End If
End If
If MenuItems(tI).Style = 2 And tI <> HoverItem Then
HoverY = Y
HoverItem = tI
Timer2.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Function OverItem(Y As Single) As Integer
Dim i, Item As clsMenuItem, Y2, LastParent As clsMenuItem
Y2 = 1
i = LBound(MenuItems)
Do
  Set Item = MenuItems(i)
  Select Case Item.Style
  Case 0 'header -22
    Set LastParent = Item
    If Item.Visible = True Then
      Y2 = Y2 + 22
      If Y < Y2 Then Exit Do
    End If
  Case 1 'sub header -18
    Set LastParent = Item
    If Item.Visible = True Then
      Y2 = Y2 + 18
      If Y < Y2 Then Exit Do
    End If
  Case 2 'popout item -18
    If LastParent Is Nothing Then GoTo DoWithSkipage2
    If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage2 Else GoTo Arg2
DoWithSkipage2:
    If Item.Visible = True Then
      Y2 = Y2 + 19
      If Y < Y2 Then Exit Do
    End If
Arg2:
  Case 3 'normal item -18
    If LastParent Is Nothing Then GoTo DoWithSkipage1
    If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage1 Else GoTo Arg1
DoWithSkipage1:
    If Item.Visible = True Then
      Y2 = Y2 + 19
      If Y < Y2 Then Exit Do
    End If
Arg1:
  End Select
  i = i + 1
Loop Until i >= UBound(MenuItems)
Dim lV, fV
lV = LastVisible
fV = FirstVisible
If i < fV Then i = fV
If i >= lV Then i = lV
OverItem = i
End Function

Private Function LastVisible() As Integer
Dim i, Item As clsMenuItem, LastParent As clsMenuItem
i = LBound(MenuItems)
LastVisible = -1
Do
  Set Item = MenuItems(i)
  Select Case Item.Style
  Case 0 'header -22
    Set LastParent = Item
    If Item.Visible = True Then
      LastVisible = i
    End If
  Case 1 'sub header -18
    Set LastParent = Item
    If Item.Visible = True Then
      LastVisible = i
    End If
  Case 2 'popout item -18
    If LastParent Is Nothing Then GoTo DoWithSkipage2
    If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage2 Else GoTo Arg2
DoWithSkipage2:
    If Item.Visible = True Then
      LastVisible = i
    End If
Arg2:
  Case 3 'normal item -18
    If LastParent Is Nothing Then GoTo DoWithSkipage1
    If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage1 Else GoTo Arg1
DoWithSkipage1:
    If Item.Visible = True Then
      LastVisible = i
    End If
Arg1:
  End Select
  i = i + 1
Loop Until i >= UBound(MenuItems)
End Function

Private Function FirstVisible() As Integer
Dim i, Item As clsMenuItem, LastParent As clsMenuItem
i = LBound(MenuItems)
FirstVisible = -1
Do
  Set Item = MenuItems(i)
  Select Case Item.Style
  Case 0 'header -22
    Set LastParent = Item
    If Item.Visible = True Then
      FirstVisible = i
      Exit Function
    End If
  Case 1 'sub header -18
    Set LastParent = Item
    If Item.Visible = True Then
      FirstVisible = i
      Exit Function
    End If
  Case 2 'popout item -18
    If LastParent Is Nothing Then GoTo DoWithSkipage2
    If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage2 Else GoTo Arg2
DoWithSkipage2:
    If Item.Visible = True Then
      FirstVisible = i
      Exit Function
    End If
Arg2:
  Case 3 'normal item -18
    If LastParent Is Nothing Then GoTo DoWithSkipage1
    If LastParent.Visible = True And LastParent.Opened = True Then GoTo DoWithSkipage1 Else GoTo Arg1
DoWithSkipage1:
    If Item.Visible = True Then
      FirstVisible = i
      Exit Function
    End If
Arg1:
  End Select
  i = i + 1
Loop Until i >= UBound(MenuItems)
End Function

Public Function HasMenu(Item As clsMenuItem) As Boolean
On Error GoTo ErrH
Dim a
a = Item.PopOutItems.Count
HasMenu = True
Exit Function
ErrH:
HasMenu = False
Exit Function
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tI As Integer, i
tI = OverItem(Y)
Select Case MenuItems(tI).Style
Case 0 ''''''''''''''''''''''''''''''''''''
'
Case 1 ''''''''''''''''''''''''''''''''''''
MenuItems(tI).Opened = Not MenuItems(tI).Opened
Case 2 ''''''''''''''''''''''''''''''''''''
Dim tArray() As clsMenuItem
If HasMenu(MenuItems(tI)) = False Then GoTo dHm
If MenuItems(tI).PopOutItems.Count > 0 Then
ReDim Preserve tArray(1 To MenuItems(tI).PopOutItems.Count)
a = MenuItems(tI).PopOutItems.Count
i = 1
Do
Set tArray(i) = MenuItems(tI).PopOutItems(i)
i = i + 1
Loop Until i > MenuItems(tI).PopOutItems.Count
Else
dHm:
ReDim tArray(1 To 1)
Set tArray(1) = New clsMenuItem
tArray(1).Caption = "Empty"
tArray(1).Style = 0
tArray(1).Visible = True
End If
If UBound(MenuReturn) > mIndex Then
If Not MenuForms(mIndex + 1) Is Nothing Then
MenuForms(mIndex + 1).ItemClicked = False
Unload MenuForms(mIndex + 1)
Set MenuForms(mIndex + 1) = Nothing
End If
ReDim Preserve MenuReturn(0 To mIndex)
ReDim Preserve MenuForms(0 To mIndex)
End If
PopMenu tArray(), mIndex + 1, Me.Left + Me.Width - (8 * Screen.TwipsPerPixelX), Me.Top + (MenuItems(tI).LastTop * Screen.TwipsPerPixelY)
i = UBound(tArray)
Do
Set tArray(i) = Nothing
i = i - 1
Loop Until i < LBound(tArray)
MenuReturn(mIndex) = tI
If mIndex = 0 Then
MenuReturn(0) = tI
End If
ReDim tArray(0)
Timer2.Enabled = False
Timer3.Enabled = False
Case 3 ''''''''''''''''''''''''''''''''''''
MenuReturn(mIndex) = tI
'If mIndex = 0 Then
'MenuReturn(0) = tI
'Else
'i = 1
'Do
'MenuReturn(0) = MenuReturn(0) & " " & MenuReturn(i)
'i = i + 1
'Loop Until i > UBound(MenuReturn)
'End If
LastItem = tI
ItemClicked = True
Unload Me
End Select
RefreshMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
If mIndex = 0 Then
If UBound(MenuReturn) > 0 Then
Dim i
i = UBound(MenuForms)
Do
If Not MenuForms(i) Is Nothing Then
Unload MenuForms(i)
Set MenuForms(i) = Nothing
MenuReturn(i) = "-1"
End If
i = i - 1
Loop Until i < 1
'ReDim Preserve MenuForms(0 To 0)
'ReDim Preserve MenuReturn(0 To 0)
End If
Else
If ItemClicked = True Then
i = 1
Do
MenuReturn(0) = MenuReturn(0) & " " & MenuReturn(i)
i = i + 1
Loop Until i > UBound(MenuReturn)
Unload MenuForms(0)
Set MenuForms(0) = Nothing
End If
End If
If UBound(MenuForms) >= mIndex Then Set MenuForms(mIndex) = Nothing
End Sub

Private Sub Timer1_Timer()
DoEvents
Dim aW As Long
aW = GetActiveWindow
If aW <> Me.hwnd Then
Dim i As Integer, j As Boolean
If mIndex < UBound(MenuReturn) Then
j = True
i = mIndex + 1
Do
If Not MenuForms(i) Is Nothing Then If aW = MenuForms(i).hwnd Then j = False
i = i + 1
Loop Until i > UBound(MenuReturn)
End If
If j = True Then Unload Me
End If
DoEvents
End Sub

Private Function PercentColor(Percent As Long, C1 As Long, C2 As Long) As Long
Dim r, g, b
r = ((RGBRed(C1) * (255 - Percent)) + (RGBRed(C2) * (Percent))) / 255
g = ((RGBGreen(C1) * (255 - Percent)) + (RGBGreen(C2) * (Percent))) / 255
b = ((RGBBlue(C1) * (255 - Percent)) + (RGBBlue(C2) * (Percent))) / 255
PercentColor = RGB(r, g, b)
End Function

Private Sub DrawHGrad(Area As RECT, C1 As Long, C2 As Long)
Dim uDC, i, tC, tP As Long
uDC = Me.hdc
i = Area.Left
Do
If i <> 0 And i <> Area.Right Then
tP = ((i - Area.Left) / (Area.Right - Area.Left)) * 100
tC = PercentColor(tP * 2, C1, C2)
ElseIf i = 0 Then
tC = C1
Else
tC = C2
End If
SetPixel uDC, i, Area.Top, tC
i = i + 1
Loop Until i > Area.Right
If Area.Bottom - Area.Top > 1 Then
StretchBlt uDC, Area.Left, Area.Top + 1, Area.Right - Area.Left, Area.Bottom - Area.Top - 1, uDC, Area.Left, Area.Top, Area.Right - Area.Left, 1, vbSrcCopy
End If
End Sub

Private Sub DrawRightArrow(uDC, X, Y, Optional aFC As Long = -1)
If aFC < 0 Then aFC = GetSetting(App.Title, "/", "ArrowFillColor", RGB(0, 0, 255))

SetPixel uDC, X + 1, Y, aFC
SetPixel uDC, X + 1, Y + 1, aFC
SetPixel uDC, X + 1, Y + 2, aFC
SetPixel uDC, X + 1, Y + 3, aFC
SetPixel uDC, X + 1, Y + 4, aFC
SetPixel uDC, X + 1, Y + 5, aFC
SetPixel uDC, X + 1, Y + 6, aFC

SetPixel uDC, X + 2, Y + 1, aFC
SetPixel uDC, X + 2, Y + 2, aFC
SetPixel uDC, X + 2, Y + 3, aFC
SetPixel uDC, X + 2, Y + 4, aFC
SetPixel uDC, X + 2, Y + 5, aFC

SetPixel uDC, X + 3, Y + 2, aFC
SetPixel uDC, X + 3, Y + 3, aFC
SetPixel uDC, X + 3, Y + 4, aFC

SetPixel uDC, X + 4, Y + 3, aFC
End Sub

Private Sub DrawDownArrow(uDC, X, Y, Optional aFC As Long = -1)
If aFC < 0 Then aFC = GetSetting(App.Title, "/", "ArrowFillColor", RGB(0, 0, 255))

SetPixel uDC, X + 3, Y + 1, aFC
SetPixel uDC, X + 4, Y + 1, aFC
SetPixel uDC, X + 5, Y + 1, aFC
SetPixel uDC, X + 6, Y + 1, aFC
SetPixel uDC, X + 7, Y + 1, aFC
SetPixel uDC, X + 8, Y + 1, aFC
SetPixel uDC, X + 9, Y + 1, aFC

SetPixel uDC, X + 4, Y + 2, aFC
SetPixel uDC, X + 5, Y + 2, aFC
SetPixel uDC, X + 6, Y + 2, aFC
SetPixel uDC, X + 7, Y + 2, aFC
SetPixel uDC, X + 8, Y + 2, aFC

SetPixel uDC, X + 5, Y + 3, aFC
SetPixel uDC, X + 6, Y + 3, aFC
SetPixel uDC, X + 7, Y + 3, aFC

SetPixel uDC, X + 6, Y + 4, aFC
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Form_MouseUp 1, 0, 0, HoverY
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
If LastItem <> HoverItem Then 'If MenuItems(LastItem).Style <> 2 Then
If UBound(MenuForms) > mIndex Then
Dim i
i = UBound(MenuForms)
Do
If Not MenuForms(i) Is Nothing Then
Unload MenuForms(i)
Set MenuForms(i) = Nothing
End If
i = i - 1
Loop Until i <= mIndex
ReDim Preserve MenuReturn(LBound(MenuReturn) To mIndex)
ReDim Preserve MenuForms(LBound(MenuForms) To mIndex)
End If
End If
End Sub
