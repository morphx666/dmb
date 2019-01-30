VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"CConfiguration"
Attribute VB_Ext_KEY = "Member0" ,"CConfiguration"
Option Explicit

'local variable to hold collection
Private mCol As Collection

Public Function Add(Key As String, Title As String, Description As String, RootWeb As String, CompiledPath As String, ImagesPath As String, OptmizePaths As Boolean, HotSpotsFile As String, Frames As CFramesConfiguration, LocalInfo4RemoteConfig As String, Optional sKey As String) As CConfiguration
    'create a new object
    Dim objNewMember As CConfiguration
    Set objNewMember = New CConfiguration


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
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As CConfiguration
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

