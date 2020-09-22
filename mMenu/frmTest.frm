VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Created By Cory J. Geesaman - cory@geesaman.com"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5805
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H004B4239&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   359
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H004B4239&
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   2
         Top             =   60
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H004B4239&
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H004B4239&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   3
      Top             =   120
      Width           =   5535
   End
   Begin VB.Image pAbout 
      Height          =   240
      Left            =   1080
      Picture         =   "frmTest.frx":1042
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pExit 
      Height          =   240
      Left            =   2040
      Picture         =   "frmTest.frx":15CC
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pFile 
      Height          =   240
      Left            =   1800
      Picture         =   "frmTest.frx":1B56
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pSave 
      Height          =   240
      Left            =   1560
      Picture         =   "frmTest.frx":1EE0
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pOpen 
      Height          =   240
      Left            =   1320
      Picture         =   "frmTest.frx":246A
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image pNew 
      Height          =   240
      Left            =   1080
      Picture         =   "frmTest.frx":29F4
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   $"frmTest.frx":2F7E
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   5520
      Picture         =   "frmTest.frx":30CB
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
Dim a
For Each a In MenuForms()
If Not a Is Nothing Then
Unload a
Set a = Nothing
End If
Next a
End Sub

Private Sub Form_Load()
ReDim MenuForms(0 To 0)
ReDim MenuReturn(0 To 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click(Index As Integer)
Dim a, aW As Long
aW = GetActiveWindow
If Me.hwnd = aW Then GoTo ItsKoo
For Each a In MenuForms()
If Not a Is Nothing Then
If a.hwnd = aW Then GoTo ItsKoo
End If
Next a
Exit Sub
ItsKoo:
Dim r As RECT, MenuData() As clsMenuItem, tItem As clsMenuItem
GetWindowRect Picture1.hwnd, r
Select Case Index
Case 0 'file
ReDim MenuData(0 To 5) 'need to dim this 1 over the actual number used, starting at 0
'the reason for the dimming 1 over is because i was getting some weird bugs the first
'few times i was running this so i had the mPopUpMenu refresh/redraw thing chop off the
'last item in the array, and i am just too lazy to fix it cause it works in my app now
Set MenuData(0) = New clsMenuItem
MenuData(0).Caption = "New"
Set MenuData(0).Icon = pNew.Picture
MenuData(0).Opened = True
MenuData(0).Style = 3
MenuData(0).Visible = True
Set MenuData(1) = New clsMenuItem
MenuData(1).Caption = "Open"
Set MenuData(1).Icon = pOpen.Picture
MenuData(1).Opened = True
MenuData(1).Style = 3
MenuData(1).Visible = True
Set MenuData(2) = New clsMenuItem
MenuData(2).Caption = "Save"
Set MenuData(2).Icon = pSave.Picture
MenuData(2).Opened = True
MenuData(2).Style = 3
MenuData(2).Visible = True
Set MenuData(3) = New clsMenuItem
MenuData(3).Caption = "Recent Files"
Set MenuData(3).Icon = pFile.Picture
MenuData(3).Opened = True
MenuData(3).Style = 2
MenuData(3).Visible = True
  Set tItem = New clsMenuItem
  tItem.Caption = "File 1"
  Set tItem.Icon = pFile.Picture
  tItem.Opened = True
  tItem.Style = 3
  tItem.Visible = True
  MenuData(3).PopOutItems.Add tItem
  Set tItem = New clsMenuItem
  tItem.Caption = "File 2"
  Set tItem.Icon = pFile.Picture
  tItem.Opened = True
  tItem.Style = 3
  tItem.Visible = True
  MenuData(3).PopOutItems.Add tItem
  Set tItem = New clsMenuItem
  tItem.Caption = "File 3"
  Set tItem.Icon = pFile.Picture
  tItem.Opened = True
  tItem.Style = 3
  tItem.Visible = True
  MenuData(3).PopOutItems.Add tItem
  MenuData(3).PopOutItems.Add tItem
Set MenuData(4) = New clsMenuItem
MenuData(4).Caption = "Exit"
Set MenuData(4).Icon = pExit.Picture
MenuData(4).Opened = True
MenuData(4).Style = 3
MenuData(4).Visible = True
Case 1 'help
ReDim MenuData(0 To 1) 'need to dim this 1 over the actual number used, starting at 0
'the reason for the dimming 1 over is because i was getting some weird bugs the first
'few times i was running this so i had the mPopUpMenu refresh/redraw thing chop off the
'last item in the array, and i am just too lazy to fix it cause it works in my app now
Set MenuData(0) = New clsMenuItem
MenuData(0).Caption = "About"
Set MenuData(0).Icon = pAbout.Picture
MenuData(0).Opened = True
MenuData(0).Style = 3
MenuData(0).Visible = True
Case 2 'this is not a real thing, just for right-clicks
ReDim MenuData(0 To 4) 'need to dim this 1 over the actual number used, starting at 0
'the reason for the dimming 1 over is because i was getting some weird bugs the first
'few times i was running this so i had the mPopUpMenu refresh/redraw thing chop off the
'last item in the array, and i am just too lazy to fix it cause it works in my app now
Set MenuData(0) = New clsMenuItem
MenuData(0).Caption = "You Right-Clicked!"
Set MenuData(0).Icon = pNew.Picture
MenuData(0).Opened = True
MenuData(0).Style = 0
MenuData(0).Visible = True
Set MenuData(1) = New clsMenuItem
MenuData(1).Caption = "This"
Set MenuData(1).Icon = pNew.Picture
MenuData(1).Opened = True
MenuData(1).Style = 2
MenuData(1).Visible = True
  Set tItem = New clsMenuItem
  tItem.Caption = "Is"
  Set tItem.Icon = pFile.Picture
  tItem.Opened = True
  tItem.Style = 3
  tItem.Visible = True
  MenuData(1).PopOutItems.Add tItem
  Set tItem = New clsMenuItem
  tItem.Caption = "Obviously"
  Set tItem.Icon = pFile.Picture
  tItem.Opened = True
  tItem.Style = 3
  tItem.Visible = True
  MenuData(1).PopOutItems.Add tItem
  Set tItem = New clsMenuItem
  tItem.Caption = "A"
  Set tItem.Icon = pFile.Picture
  tItem.Opened = True
  tItem.Style = 3
  tItem.Visible = True
  MenuData(1).PopOutItems.Add tItem
  MenuData(1).PopOutItems.Add tItem
Set MenuData(2) = New clsMenuItem
MenuData(2).Caption = "Right-Click"
Set MenuData(2).Icon = pNew.Picture
MenuData(2).Opened = True
MenuData(2).Style = 3
MenuData(2).Visible = True
Set MenuData(3) = New clsMenuItem
MenuData(3).Caption = "Menu"
Set MenuData(3).Icon = pNew.Picture
MenuData(3).Opened = True
MenuData(3).Style = 3
MenuData(3).Visible = True
End Select
MenuChanged = False
If Index <> 2 Then
PopMenu MenuData(), 0, (r.Left + Label1(Index).Left) * Screen.TwipsPerPixelX, (r.Top + Label1(Index).Top + Label1(Index).Height) * Screen.TwipsPerPixelY
Else
PopMenu MenuData(), 0
End If
Do
DoEvents
Loop Until MenuForms(0) Is Nothing Or MenuChanged = True
Select Case Index
Case 0
Select Case MenuReturn(0)
Case "0"
MsgBox "new"
Case "1"
MsgBox "open"
Case "2"
MsgBox "save"
Case "3 1"
MsgBox "file1"
Case "3 2"
MsgBox "file2"
Case "3 3"
MsgBox "file3"
Case "4"
MsgBox "exit"
End Select
Case 1
Select Case MenuReturn(0)
Case "0"
MsgBox "about"
End Select
Case 2
Select Case MenuReturn(0)
Case "1"
MsgBox "this"
Case "2 1"
MsgBox "is"
Case "2 2"
MsgBox "obviously"
Case "2 3"
MsgBox "a"
Case "3"
MsgBox "right-click"
Case "4"
MsgBox "menu"
End Select
End Select
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Label1_Click 2
End Sub
