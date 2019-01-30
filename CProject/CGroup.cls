VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Public Enum GroupAlignmentConstants
    gacBottomLeft = 0
    gacBottomRight = 1
    gacTopLeft = 2
    gacTopRight = 3
    gacLeftTop = 4
    gacLeftBottom = 5
    gacRightTop = 6
    gacRightBottom = 7
    gacTopCenter = 8
    gacRightCenter = 9
    gacBottomCenter = 10
    gacLeftCenter = 11
End Enum

