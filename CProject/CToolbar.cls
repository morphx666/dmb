VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarTitle As String 'local copy
Private mvarFollowHScroll As Boolean 'local copy
Private mvarFollowVScroll As Boolean 'local copy
Public Enum ToolbarAlignmentConstants
    tbacTopLeft = 0
    tbacTopCenter = 1
    tbacTopRight = 2
    tbacCenterLeft = 3
    tbacCenterCenter = 4
    tbacCenterRight = 5
    tbacBottomLeft = 6
    tbacBottomCenter = 7
    tbacBottomRight = 8
End Enum
Public Enum ToolbarStyleConstants
    tscHorizonal = 0
    tscVertical = 1
End Enum
Public Enum ToolbarSpanningConstants
    tscAuto = 0
    tscFull = 1
End Enum
Private mvarAlignment As ToolbarAlignmentConstants 'local copy
Private mvarAttachTo As String 'local copy
Private mvarAttachToAlignment As GroupAlignmentConstants 'local copy
Private mvarMarginH As Integer 'local copy
Private mvarMarginV As Integer 'local copy
Private mvarStyle As ToolbarStyleConstants 'local copy
Private mvarSpanning As ToolbarSpanningConstants 'local copy
Private mvarBorder As Integer 'local copy
Public Property Let Border(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Border = 5
    mvarBorder = vData
End Property


Public Property Get Border() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Border
    Border = mvarBorder
End Property



Public Property Set Spanning(ByVal vData As ToolbarSpanningConstants)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Spanning = Form1
    Set mvarSpanning = vData
End Property


Public Property Get Spanning() As ToolbarSpanningConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Spanning
    Set Spanning = mvarSpanning
End Property



Public Property Set Style(ByVal vData As ToolbarStyleConstants)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Style = Form1
    Set mvarStyle = vData
End Property


Public Property Get Style() As ToolbarStyleConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Style
    Set Style = mvarStyle
End Property



Public Property Let MarginV(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MarginV = 5
    mvarMarginV = vData
End Property


Public Property Get MarginV() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MarginV
    MarginV = mvarMarginV
End Property



Public Property Let MarginH(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MarginH = 5
    mvarMarginH = vData
End Property


Public Property Get MarginH() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MarginH
    MarginH = mvarMarginH
End Property



Public Property Set AttachToAlignment(ByVal vData As GroupAlignmentConstants)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AttachToAlignment = Form1
    Set mvarAttachToAlignment = vData
End Property


Public Property Get AttachToAlignment() As GroupAlignmentConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttachToAlignment
    Set AttachToAlignment = mvarAttachToAlignment
End Property



Public Property Let AttachTo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttachTo = 5
    mvarAttachTo = vData
End Property


Public Property Get AttachTo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttachTo
    AttachTo = mvarAttachTo
End Property



Public Property Let Alignment(ByVal vData As ToolbarAlignmentConstants)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Alignment = 5
    mvarAlignment = vData
End Property


Public Property Get Alignment() As ToolbarAlignmentConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Alignment
    Set Alignment = mvarAlignment
End Property






Public Property Let FollowVScroll(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FollowVScroll = 5
    mvarFollowVScroll = vData
End Property


Public Property Get FollowVScroll() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FollowVScroll
    FollowVScroll = mvarFollowVScroll
End Property



Public Property Let FollowHScroll(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FollowHScroll = 5
    mvarFollowHScroll = vData
End Property


Public Property Get FollowHScroll() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FollowHScroll
    FollowHScroll = mvarFollowHScroll
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



