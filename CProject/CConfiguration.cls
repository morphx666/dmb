VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Public Enum ConfigTypeContants
    ctcLocal
    ctcRemote
    ctcCDROM
End Enum

Private mvarTitle As String 'local copy
Private mvarDescription As String 'local copy
Private mvarRootWeb As String 'local copy
Private mvarCompiledPath As String 'local copy
Private mvarImagesPath As String 'local copy
Private mvarOptmizePaths As Boolean 'local copy
Private mvarHotSpotsFile As String 'local copy
Private mvarFrames As CFramesConfiguration 'local copy
Private mvarLocalInfo4RemoteConfig As String 'local copy
Public Property Let LocalInfo4RemoteConfig(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LocalInfo4RemoteConfig = 5
    mvarLocalInfo4RemoteConfig = vData
End Property


Public Property Get LocalInfo4RemoteConfig() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LocalInfo4RemoteConfig
    LocalInfo4RemoteConfig = mvarLocalInfo4RemoteConfig
End Property



Public Property Set Frames(ByVal vData As CFramesConfiguration)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Frames = Form1
    Set mvarFrames = vData
End Property


Public Property Get Frames() As CFramesConfiguration
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Frames
    Set Frames = mvarFrames
End Property



Public Property Let HotSpotsFile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HotSpotsFile = 5
    mvarHotSpotsFile = vData
End Property


Public Property Get HotSpotsFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HotSpotsFile
    HotSpotsFile = mvarHotSpotsFile
End Property



Public Property Let OptmizePaths(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OptmizePaths = 5
    mvarOptmizePaths = vData
End Property


Public Property Get OptmizePaths() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OptmizePaths
    OptmizePaths = mvarOptmizePaths
End Property



Public Property Let ImagesPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ImagesPath = 5
    mvarImagesPath = vData
End Property


Public Property Get ImagesPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ImagesPath
    ImagesPath = mvarImagesPath
End Property



Public Property Let CompiledPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CompiledPath = 5
    mvarCompiledPath = vData
End Property


Public Property Get CompiledPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompiledPath
    CompiledPath = mvarCompiledPath
End Property



Public Property Let RootWeb(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootWeb = 5
    mvarRootWeb = vData
End Property


Public Property Get RootWeb() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootWeb
    RootWeb = mvarRootWeb
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



