VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Toolbars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CToolbar"
Attribute VB_Ext_KEY = "Member0" ,"CToolbar"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, Title As String, FollowHScroll As Boolean, FollowVScroll As Boolean, Alignment As ToolbarAlignmentConstants, AttachTo As String, AttachToAlignment As GroupAlignmentConstants, MarginH As Integer, MarginV As Integer, Style As ToolbarStyleConstants, Spanning As ToolbarSpanningConstants, Border As Integer, Optional sKey As String) As CToolbar
    'create a new object
    Dim objNewMember As CToolbar
    Set objNewMember = New CToolbar


    'set the properties passed into the method
    objNewMember.Key = Key
    objNewMember.Title = Title
    objNewMember.Description = Description
    objNewMember.RootWeb = RootWeb
    objNewMember.CompiledPath = CompiledPath
    objNewMember.ImagesPath = ImagesPath
    objNewMember.OptmizePaths = OptmizePaths
    objNewMember.HotSpotsFile = HotSpotsFile
    If IsObject(Frames) Then
        Set objNewMember.Frames = Frames
    Else
        objNewMember.Frames = Frames
    End If
    objNewMember.LocalInfo4RemoteConfig = LocalInfo4RemoteConfig
    objNewMember.Key = Key
    objNewMember.Title = Title
    objNewMember.FollowHScroll = FollowHScroll
    objNewMember.FollowVScroll = FollowVScroll
    If IsObject(Alignment) Then
        Set objNewMember.Alignment = Alignment
    Else
        objNewMember.Alignment = Alignment
    End If
    objNewMember.AttachTo = AttachTo
    If IsObject(AttachToAlignment) Then
        Set objNewMember.AttachToAlignment = AttachToAlignment
    Else
        objNewMember.AttachToAlignment = AttachToAlignment
    End If
    objNewMember.MarginH = MarginH
    objNewMember.MarginV = MarginV
    If IsObject(Style) Then
        Set objNewMember.Style = Style
    Else
        objNewMember.Style = Style
    End If
    If IsObject(Spanning) Then
        Set objNewMember.Spanning = Spanning
    Else
        objNewMember.Spanning = Spanning
    End If
    objNewMember.Border = Border
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CToolbar
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

