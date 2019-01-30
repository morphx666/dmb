VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Toolbars"
Option Explicit

Private mvarTitle As String 'local copy
Private mvarDescription As String 'local copy
Private mvarMinVersion As String 'local copy
Private mvarToolbars As Toolbars


Public Property Get Toolbars() As Toolbars
    If mvarToolbars Is Nothing Then
        Set mvarToolbars = New Toolbars
    End If


    Set Toolbars = mvarToolbars
End Property


Public Property Set Toolbars(vData As Toolbars)
    Set mvarToolbars = vData
End Property
Private Sub Class_Terminate()
    Set mvarToolbars = Nothing
End Sub



Public Property Let MinVersion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinVersion = 5
    mvarMinVersion = vData
End Property


Public Property Get MinVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinVersion
    MinVersion = mvarMinVersion
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property


Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
End Property



