VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCaption As String 'local copy
Private mvarStyle As Integer 'local copy
Private mvarPopOutItems As Collection 'local copy
'local variable(s) to hold property value(s)
Private mvarVisible As Boolean 'local copy
Private mvarOpened As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarIcon As Picture 'local copy
'local variable(s) to hold property value(s)
Private mvarSelected As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarLastTop As Long 'local copy
Public Property Let LastTop(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastTop = 5
    mvarLastTop = vData
End Property


Public Property Get LastTop() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastTop
    LastTop = mvarLastTop
End Property



Public Property Let Selected(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Selected = 5
    mvarSelected = vData
End Property


Public Property Get Selected() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Selected
    Selected = mvarSelected
End Property



Public Property Set Icon(ByVal vData As Picture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Icon = Form1
    Set mvarIcon = vData
End Property


Public Property Get Icon() As Picture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Icon
    Set Icon = mvarIcon
End Property



Public Property Let Opened(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Opened = 5
    mvarOpened = vData
End Property


Public Property Get Opened() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Opened
    Opened = mvarOpened
End Property



Public Property Let Visible(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Visible = 5
    mvarVisible = vData
End Property


Public Property Get Visible() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
End Property




Public Property Set PopOutItems(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PopOutItems = Form1
    Set mvarPopOutItems = vData
End Property

Public Property Get PopOutItems() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PopOutItems
    Set PopOutItems = mvarPopOutItems
End Property

Public Property Let Style(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Border = 5
    mvarStyle = vData
End Property

Public Property Get Style() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Border
    Style = mvarStyle
End Property

Public Property Let Caption(ByVal vData As String)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Caption = Form1
    mvarCaption = vData
End Property

Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property

Private Sub Class_Initialize()
Set mvarPopOutItems = New Collection
End Sub
