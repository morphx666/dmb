VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CAddIn"
Attribute VB_Ext_KEY = "Member1" ,"CFramesConfiguration"
Attribute VB_Ext_KEY = "Member2" ,"Configurations"
Attribute VB_Ext_KEY = "Member3" ,"Groups"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarTitle As String 'local copy
'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
Private mvarHasChanged As Boolean 'local copy
Private mvarOriginalFileName As String 'local copy
Private mvarFX As Integer 'local copy
Public Enum CodeOptimizationConstants
    cocDEBUG = 0
    cocNormal = 1
    cocAggressive = 2
End Enum
Private mvarCodeOptimization As CodeOptimizationConstants 'local copy
Private mvarGenDynAPI As Boolean 'local copy
Private mvarConfigurations As Configurations
Private mvarCFramesConfiguration As CFramesConfiguration
Private mvarDefaultConfig As CConfiguration 'local copy
Private mvarCAddIn As CAddIn
Private mvarGroups As Groups


Public Property Get Groups() As Groups
    If mvarGroups Is Nothing Then
        Set mvarGroups = New Groups
    End If


    Set Groups = mvarGroups
End Property


Public Property Set Groups(vData As Groups)
    Set mvarGroups = vData
End Property



Public Property Get CAddIn() As CAddIn
    Set CAddIn = mvarCAddIn
End Property


Public Property Set CAddIn(vData As CAddIn)
    Set mvarCAddIn = vData
End Property

Public Property Set DefaultConfig(ByVal vData As CConfiguration)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DefaultConfig = Form1
    Set mvarDefaultConfig = vData
End Property


Public Property Get DefaultConfig() As CConfiguration
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultConfig
    Set DefaultConfig = mvarDefaultConfig
End Property





Private Sub Class_Initialize()
    'create the mCFramesConfiguration object when the CProject class is created
    Set mvarCFramesConfiguration = New CFramesConfiguration
  'create the mCAddIn object when the CProject class is created
  Set mvarCAddIn = New CAddIn
End Sub


Public Property Get CFramesConfiguration() As CFramesConfiguration
    Set CFramesConfiguration = mvarCFramesConfiguration
End Property


Public Property Set CFramesConfiguration(vData As CFramesConfiguration)
    Set mvarCFramesConfiguration = vData
End Property



Public Property Get Configurations() As Configurations
    If mvarConfigurations Is Nothing Then
        Set mvarConfigurations = New Configurations
    End If


    Set Configurations = mvarConfigurations
End Property


Public Property Set Configurations(vData As Configurations)
    Set mvarConfigurations = vData
End Property
Private Sub Class_Terminate()
  Set mvarGroups = Nothing
  Set mvarCAddIn = Nothing
  Set mvarCFramesConfiguration = Nothing
    Set mvarConfigurations = Nothing
End Sub



Public Property Let GenDynAPI(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GenDynAPI = 5
    mvarGenDynAPI = vData
End Property


Public Property Get GenDynAPI() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GenDynAPI
    GenDynAPI = mvarGenDynAPI
End Property



Public Property Let CodeOptimization(ByVal vData As CodeOptimizationConstants)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeOptimization = 5
    mvarCodeOptimization = vData
End Property


Public Property Get CodeOptimization() As CodeOptimizationConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeOptimization
    Set CodeOptimization = mvarCodeOptimization
End Property




Public Property Let FX(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FX = 5
    mvarFX = vData
End Property


Public Property Get FX() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FX
    FX = mvarFX
End Property



Public Property Let OriginalFileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OriginalFileName = 5
    mvarOriginalFileName = vData
End Property


Public Property Get OriginalFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OriginalFileName
    OriginalFileName = mvarOriginalFileName
End Property



Public Property Let HasChanged(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HasChanged = 5
    mvarHasChanged = vData
End Property


Public Property Get HasChanged() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HasChanged
    HasChanged = mvarHasChanged
End Property



Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
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



